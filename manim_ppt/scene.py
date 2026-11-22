"""``PPTScene``: a drop-in replacement for ``manim.Scene`` that also writes a
PowerPoint deck made of native, editable shapes animated with Morph."""

from __future__ import annotations

import math
from pathlib import Path
from typing import Any

from manim import (
    AnimationGroup,
    Homotopy,
    MoveAlongPath,
    Rotating,
    Scene,
    Succession,
    Transform,
    Wait,
    config,
    logger,
)
from manim.utils.rate_functions import (
    there_and_back,
    there_and_back_with_pause,
    wiggle,
)

from .capture import MobjectNamer, capture
from .model import Keyframe, Snapshot, Timeline
from .writer import _shape_distance, write_pptx

_RETURNING_RATE_FUNCS = {there_and_back, there_and_back_with_pause}


def default_samples(anim: Any) -> int:
    """How many Morph steps an animation needs to read correctly in PowerPoint.

    Morph moves every point in a straight line from one slide to the next, so
    one step is right for most animations (moves, fades, recolors, transforms).
    Curved motion and animations that return to their start need more.
    """
    if isinstance(anim, Succession):
        return sum(default_samples(a) for a in anim.animations)
    if isinstance(anim, AnimationGroup):
        n = max((default_samples(a) for a in anim.animations), default=1)
        if getattr(anim, "lag_ratio", 0) > 0 and len(anim.animations) > 1:
            n = max(n, min(len(anim.animations), 6))
        return n
    rate = getattr(anim, "rate_func", None)
    if rate is wiggle:
        return 8
    if rate in _RETURNING_RATE_FUNCS:
        return 2
    if isinstance(anim, Rotating):
        return max(1, math.ceil(abs(anim.angle) / (math.pi / 4)))
    if isinstance(anim, Transform) and getattr(anim, "path_arc", 0):
        return max(1, math.ceil(abs(anim.path_arc) / (math.pi / 4)))
    if isinstance(anim, (MoveAlongPath, Homotopy)):
        return 6
    if isinstance(anim, Wait):
        return max(1, math.ceil(anim.run_time * 4))
    return 1


class PPTScene(Scene):
    """Use exactly like ``Scene``; a ``.pptx`` is written next to the video.

    Slide flow
    ----------
    * By default every ``self.play(...)`` waits for a click, like a build in
      PowerPoint. ``self.wait()`` calls are ignored in this mode.
    * As soon as you call ``self.next_slide()`` anywhere, the deck switches to
      marker mode: everything between two ``next_slide()`` calls plays on one
      click, and ``self.wait(t)`` becomes a timed pause.

    Class attributes
    ----------------
    ppt_output:
        Output path. Defaults to ``<media_dir>/pptx/<SceneName>.pptx``.
    ppt_mode:
        ``"auto"`` (default), ``"click"`` or ``"markers"``.
    ppt_max_samples:
        Upper bound on Morph steps generated for one ``play()``.
    ppt_updater_samples_per_second:
        Morph steps per second used while updaters are active.
    """

    ppt_output: str | None = None
    ppt_mode: str = "auto"
    ppt_max_samples: int = 12
    ppt_updater_samples_per_second: float = 4.0

    def __init__(self, *args: Any, **kwargs: Any) -> None:
        super().__init__(*args, **kwargs)
        # Cached animations are skipped wholesale, which would hide their
        # intermediate states from us.
        config.disable_caching = True
        self._ppt_timeline = Timeline(
            frame_width=config.frame_width, frame_height=config.frame_height
        )
        self._ppt_namer = MobjectNamer()
        self._ppt_markers_used = False
        self._ppt_last_fp: str | None = None
        self._ppt_pre_fp: str | None = None
        self._ppt_times: list[float] = []
        self._ppt_last_t = 0.0
        self._ppt_override: int | None = None
        self._ppt_pending_start = False
        self._ppt_notes: str | None = None
        self._ppt_sampling = False
        self._ppt_precleanup: Snapshot | None = None
        self._ppt_initial_marker = False

    # ------------------------------------------------------------------ API

    def next_slide(self, notes: str | None = None) -> None:
        """Pause here until the presenter clicks.

        ``notes`` become the speaker notes of the slides that follow.
        """
        self._ppt_markers_used = True
        if self._ppt_timeline.keyframes:
            self._ppt_timeline.keyframes[-1].marker = True
        else:
            self._ppt_initial_marker = True
        self._ppt_notes = notes
        self._ppt_pending_start = True

    def play(self, *args: Any, ppt_samples: int | None = None, **kwargs: Any) -> None:
        """Same as ``Scene.play``. ``ppt_samples`` forces the number of Morph
        steps used for this call (more steps = closer to manim's motion)."""
        self._ppt_override = ppt_samples
        self._ppt_pre_fp = self._ppt_snapshot().fingerprint()
        self._ppt_pending_start = True
        super().play(*args, **kwargs)

        if self._ppt_is_static_wait():
            self._ppt_on_static_wait(self.duration)
        elif self._ppt_sampling:
            if self._ppt_precleanup is not None:
                self._ppt_inherit_names(self._ppt_precleanup, self._ppt_snapshot())
                self._ppt_precleanup = None
            final = self._ppt_snapshot()
            self._ppt_append(final, self.duration - self._ppt_last_t)
            self._ppt_last_fp = final.fingerprint()
        self._ppt_sampling = False
        self._ppt_times = []

    # ------------------------------------------------------- manim hooks

    def begin_animations(self) -> None:
        super().begin_animations()
        if self._ppt_is_static_wait():
            return
        tl = self._ppt_timeline
        snap0 = self._ppt_snapshot()
        if not tl.keyframes:
            tl.keyframes.append(
                Keyframe(snap0, marker=self._ppt_initial_marker, notes=self._ppt_notes)
            )
        elif self._ppt_pre_fp == self._ppt_last_fp:
            # Nothing changed since the last slide, so swap in the state at
            # alpha=0. It looks identical but carries the animation's start
            # geometry (e.g. a point for GrowFromCenter, aligned submobjects
            # for Transform), which gives Morph something to tween from.
            tl.keyframes[-1].snapshot = snap0
        else:
            # Something was add()ed / remove()d between plays: show it with a
            # hard cut, on the same click as the animation.
            self._ppt_append(snap0, 0.0)

        n = self._ppt_override or max(
            (default_samples(a) for a in (self.animations or [])), default=1
        )
        if self._ppt_override is None and self._ppt_has_updaters():
            # Updaters (always_redraw, ValueTracker-driven geometry...) can move
            # things along arbitrary paths; sample densely enough to follow.
            n = max(n, math.ceil(self.duration * self.ppt_updater_samples_per_second))
        n = max(1, min(int(n), self.ppt_max_samples))
        self._ppt_times = [self.duration * k / n for k in range(1, n)]
        self._ppt_last_t = 0.0
        self._ppt_sampling = True
        self._ppt_watch_cleanup()

    def update_to_time(self, t: float) -> None:
        super().update_to_time(t)
        while self._ppt_times and t >= self._ppt_times[0] - 1e-9:
            self._ppt_times.pop(0)
            self._ppt_append(self._ppt_snapshot(), t - self._ppt_last_t)
            self._ppt_last_t = t

    def render(self, preview: bool = False) -> bool:
        result = super().render(preview)
        self.ppt_export()
        return result

    # ------------------------------------------------------------ export

    def ppt_timeline(self) -> Timeline:
        tl = self._ppt_timeline
        final = self._ppt_snapshot()
        if not tl.keyframes:
            tl.keyframes.append(Keyframe(final))
        elif final.fingerprint() != self._ppt_last_fp:
            self._ppt_pending_start = True
            self._ppt_append(final, 0.0)
            self._ppt_last_fp = final.fingerprint()
        mode = self.ppt_mode
        if mode == "auto":
            mode = "markers" if self._ppt_markers_used else "click"
        tl.mode = mode
        return tl

    def ppt_export(self, path: str | Path | None = None) -> Path:
        path = Path(
            path
            or self.ppt_output
            or Path(config.media_dir) / "pptx" / f"{type(self).__name__}.pptx"
        )
        tl = self.ppt_timeline()
        out = write_pptx(tl, path)
        logger.info(
            "manim-ppt: wrote %(n)d slides to %(path)s",
            {"n": len(tl.keyframes), "path": str(out)},
        )
        print(f"manim-ppt: wrote {len(tl.keyframes)} slides -> {out}")
        return out

    # ---------------------------------------------------------- internals

    def _ppt_snapshot(self) -> Snapshot:
        return capture(self, self._ppt_namer)

    def _ppt_watch_cleanup(self) -> None:
        """Snapshot the scene right before the first animation cleans up.

        At that instant every animation has reached alpha=1 but mobjects that
        ``ReplacementTransform`` & co. swap out are still in the scene, which
        lets us hand their names over to the mobjects that replace them.
        """
        if not self.animations:
            return
        anim = self.animations[0]
        original = anim.clean_up_from_scene

        def clean_up_from_scene(scene: Scene) -> Any:
            anim.clean_up_from_scene = original
            self._ppt_precleanup = self._ppt_snapshot()
            return original(scene)

        anim.clean_up_from_scene = clean_up_from_scene

    def _ppt_inherit_names(self, before: Snapshot, after: Snapshot) -> None:
        before_keys = {s.key for s in before.shapes}
        after_keys = {s.key for s in after.shapes}
        new = [s for s in after.shapes if s.key not in before_keys]
        gone = [s for s in before.shapes if s.key not in after_keys]
        if not new or not gone:
            return
        tol = 0.01 * config.frame_width
        pairs = []
        for i, n in enumerate(new):
            for j, g in enumerate(gone):
                d = _shape_distance(n, g)
                if d is not None and d < tol:
                    pairs.append((d, i, j))
        pairs.sort()
        used_n: set[int] = set()
        used_g: set[int] = set()
        for _, i, j in pairs:
            if i in used_n or j in used_g:
                continue
            mob = self._ppt_namer.mobject(new[i].key)
            if mob is not None:
                self._ppt_namer.transfer(gone[j].key, mob)
            used_n.add(i)
            used_g.add(j)

    def _ppt_has_updaters(self) -> bool:
        if getattr(self, "updaters", None):
            return True
        return any(mob.get_family_updaters() for mob in self.mobjects)

    def _ppt_is_static_wait(self) -> bool:
        anims = getattr(self, "animations", None) or []
        return (
            len(anims) == 1
            and isinstance(anims[0], Wait)
            and bool(getattr(anims[0], "is_static_wait", False))
        )

    def _ppt_on_static_wait(self, duration: float) -> None:
        tl = self._ppt_timeline
        if not tl.keyframes or self._ppt_pre_fp != self._ppt_last_fp:
            snap = self._ppt_snapshot()
            self._ppt_append(snap, 0.0)
            self._ppt_last_fp = snap.fingerprint()
        tl.keyframes[-1].hold_ms += int(round(duration * 1000))

    def _ppt_append(self, snap: Snapshot, transition_s: float) -> None:
        tl = self._ppt_timeline
        ms = max(0, int(round(transition_s * 1000)))
        if tl.keyframes and snap.fingerprint() == tl.keyframes[-1].snapshot.fingerprint():
            # No visible change: just extend the time spent on the last slide.
            tl.keyframes[-1].hold_ms += ms
            return
        tl.keyframes.append(
            Keyframe(
                snap,
                transition_ms=ms,
                starts_play=self._ppt_pending_start,
                notes=self._ppt_notes,
            )
        )
        self._ppt_pending_start = False
