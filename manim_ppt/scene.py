"""``PPTScene``: a drop-in replacement for ``manim.Scene`` that also writes a
PowerPoint deck made of native, editable shapes animated with Morph."""

from __future__ import annotations

import math
from collections.abc import Iterable, Iterator
from itertools import pairwise
from pathlib import Path
from typing import Any

from manim import (
    AnimationGroup,
    DrawBorderThenFill,
    Homotopy,
    MoveAlongPath,
    Rotating,
    Scene,
    ShowIncreasingSubsets,
    ShowPartial,
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
from .model import Keyframe, PathShape, Shape, Snapshot, Timeline, bounds_distance
from .writer import write_deck

_RETURNING_RATE_FUNCS = {there_and_back, there_and_back_with_pause}
# Animations that trace a mobject into (or out of) existence. Morph cannot
# draw a stroke, so these become fades rather than a zoom out of a dot.
_DRAWING_ANIMATIONS = (ShowPartial, DrawBorderThenFill)


def default_samples(anim: Any) -> int:
    """How many Morph steps an animation needs to read correctly in PowerPoint.

    Morph moves every shape in a straight line from one slide to the next, so
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
    if isinstance(anim, ShowIncreasingSubsets):
        return max(1, min(len(getattr(anim, "all_submobs", ())), 8))
    if isinstance(anim, Wait):
        return max(1, math.ceil(anim.run_time * 4))
    return 1


def _sample_fractions(animations: list[Any], n: int) -> list[float]:
    """Where in a play() to take its ``n - 1`` intermediate keyframes.

    Evenly spaced in *progress* when the play has one easing curve, so that a
    half turn split in four is four 45 degree steps (each taking as long as
    manim spends on it) rather than two small steps and two large ones.
    Otherwise evenly spaced in time.
    """
    fractions = [k / n for k in range(1, n)]
    if not fractions or any(isinstance(a, AnimationGroup) for a in animations):
        return fractions
    rates = {id(getattr(a, "rate_func", None)) for a in animations}
    run_times = {round(float(getattr(a, "run_time", 0.0)), 6) for a in animations}
    rate = getattr(animations[0], "rate_func", None) if animations else None
    if len(rates) != 1 or len(run_times) != 1 or not callable(rate):
        return fractions
    xs = [i / 200 for i in range(201)]
    try:
        ys = [float(rate(x)) for x in xs]
    except Exception:
        return fractions
    monotonic = all(b >= a - 1e-9 for a, b in pairwise(ys))
    if not monotonic or abs(ys[0]) > 1e-6 or abs(ys[-1] - 1) > 1e-6:
        return fractions
    out = []
    for target in fractions:
        i = next(j for j, y in enumerate(ys) if y >= target)
        if i == 0:
            out.append(0.0)
            continue
        span = ys[i] - ys[i - 1]
        out.append(xs[i - 1] + (xs[i] - xs[i - 1]) * ((target - ys[i - 1]) / span if span else 0.0))
    return out


def _leaf_animations(animations: Iterable[Any]) -> Iterator[Any]:
    for anim in animations:
        if isinstance(anim, AnimationGroup):
            yield from _leaf_animations(anim.animations)
        else:
            yield anim


def _with_ghosts(current: Snapshot, earlier: Snapshot) -> Snapshot:
    """``current`` plus the shapes that ``earlier`` shows fading or shrinking
    away, each slotted in after the shape it used to follow.

    A mobject removed by ``FadeOut(m, shift=DOWN)`` or ``ShrinkToCenter(m)``
    is simply gone from the scene afterwards. Leaving its last (invisible)
    state on the slide gives Morph a destination, so the exit actually moves.
    """
    present = {s.key for s in current.shapes}
    after: dict[str | None, list[Shape]] = {}
    anchor: str | None = None
    found = False
    for shape in earlier.shapes:
        if shape.key in present:
            anchor = shape.key
        elif isinstance(shape, PathShape) and not shape.visible:
            after.setdefault(anchor, []).append(shape)
            found = True
    if not found:
        return current
    shapes = list(after.get(None, []))
    for shape in current.shapes:
        shapes.append(shape)
        shapes.extend(after.get(shape.key, []))
    return Snapshot(shapes=shapes, background=current.background)


class PPTMixin:
    """Adds PowerPoint export to any manim ``Scene`` class.

    ``PPTScene`` is this mixin applied to ``Scene``. For other scene types,
    put the mixin first::

        class Zoom(PPTMixin, MovingCameraScene): ...

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
    ppt_template:
        A ``.pptx`` to build on: its theme, slide size and background are
        used, and the animation is appended after any slides it contains.
    ppt_background:
        Paint the scene's background color onto the slides. Defaults to yes
        for a fresh deck and no when a template is used.
    ppt_native_shapes:
        Write rectangles, ellipses, lines... as PowerPoint presets and keep
        rotation as a real rotation (default). ``False`` gives freeforms only.
    ppt_group:
        Write a compound mobject (a word, an equation, axes...) that only ever
        moves as a whole as one PowerPoint group (default).
    ppt_max_samples:
        Upper bound on Morph steps generated for one ``play()``.
    ppt_updater_samples_per_second:
        Morph steps per second used while updaters are active.
    ppt_auto_export:
        Write the deck when the scene finishes rendering (default). Turn off
        to call :meth:`ppt_export` yourself.
    """

    ppt_output: str | Path | None = None
    ppt_mode: str = "auto"
    ppt_template: str | Path | None = None
    ppt_background: bool | None = None
    ppt_native_shapes: bool = True
    ppt_group: bool = True
    ppt_max_samples: int = 12
    ppt_updater_samples_per_second: float = 4.0
    ppt_auto_export: bool = True

    def __init__(self, *args: Any, **kwargs: Any) -> None:
        super().__init__(*args, **kwargs)
        if self.ppt_mode not in ("auto", "click", "markers"):
            raise ValueError(
                f"ppt_mode must be 'auto', 'click' or 'markers', not {self.ppt_mode!r}"
            )
        if "cairo" not in str(config.renderer).lower():
            logger.warning(
                "manim-ppt only supports the Cairo renderer; the deck will be incomplete "
                "with renderer=%s.",
                config.renderer,
            )
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
        self._ppt_section: str | None = None
        self._ppt_fade_keys: frozenset[str] = frozenset()
        self._ppt_sampling = False
        self._ppt_precleanup: Snapshot | None = None
        self._ppt_initial_marker = False
        self.ppt_path: Path | None = None
        self.ppt_slide_count = 0

    # ------------------------------------------------------------------ API

    def next_slide(self, notes: str | None = None, *, section: str | None = None) -> None:
        """Pause here until the presenter clicks.

        Parameters
        ----------
        notes
            Speaker notes for the slides that follow.
        section
            Start a named PowerPoint section here. Sections show up in the
            slide sorter and thumbnail pane, which keeps long decks navigable.
        """
        self._ppt_markers_used = True
        if self._ppt_timeline.keyframes:
            self._ppt_timeline.keyframes[-1].marker = True
        else:
            self._ppt_initial_marker = True
        self._ppt_notes = notes
        if section is not None:
            self._ppt_section = section
        self._ppt_pending_start = True

    def play(self, *args: Any, ppt_samples: int | None = None, **kwargs: Any) -> None:
        """Same as ``Scene.play``. ``ppt_samples`` forces the number of Morph
        steps used for this call (more steps = closer to manim's motion)."""
        self._ppt_override = ppt_samples
        self._ppt_pre_fp = self._ppt_snapshot().fingerprint()
        self._ppt_pending_start = True
        super().play(*args, **kwargs)  # type: ignore[misc]

        if self._ppt_is_static_wait():
            self._ppt_on_static_wait(self.duration)  # type: ignore[attr-defined]
        elif self._ppt_sampling:
            before, self._ppt_precleanup = self._ppt_precleanup, None
            if before is not None:
                self._ppt_inherit_names(before, self._ppt_snapshot())
            final = self._ppt_snapshot()
            if before is not None:
                final = _with_ghosts(final, before)
            self._ppt_append(final, self.duration - self._ppt_last_t)  # type: ignore[attr-defined]
            self._ppt_last_fp = final.fingerprint()
        self._ppt_sampling = False
        self._ppt_times = []
        self._ppt_fade_keys = frozenset()

    # ------------------------------------------------------- manim hooks

    def begin_animations(self) -> None:
        super().begin_animations()  # type: ignore[misc]
        if self._ppt_is_static_wait():
            return
        tl = self._ppt_timeline
        snap0 = self._ppt_snapshot()
        self._ppt_fade_keys = self._ppt_drawn_keys()
        if not tl.keyframes:
            self._ppt_new_keyframe(snap0, 0, starts_play=False, marker=self._ppt_initial_marker)
        elif self._ppt_pre_fp == self._ppt_last_fp:
            # Nothing changed since the last slide. Keep the state at alpha=0
            # next to it: it looks identical but carries the animation's start
            # geometry (e.g. a point for GrowFromCenter, aligned submobjects
            # for Transform), which gives Morph something to tween from.
            last = tl.keyframes[-1]
            last.departure = snap0
            last.fade_keys |= self._ppt_fade_keys
        else:
            # Something was add()ed / remove()d between plays: show it with a
            # hard cut, on the same click as the animation.
            self._ppt_append(snap0, 0.0)

        animations = self.animations or []  # type: ignore[attr-defined]
        duration = self.duration  # type: ignore[attr-defined]
        n = self._ppt_override or max((default_samples(a) for a in animations), default=1)
        if self._ppt_override is None and self._ppt_has_updaters():
            # Updaters (always_redraw, ValueTracker-driven geometry...) can move
            # things along arbitrary paths; sample densely enough to follow.
            n = max(n, math.ceil(duration * self.ppt_updater_samples_per_second))
        n = max(1, min(int(n), self.ppt_max_samples))
        self._ppt_times = [duration * f for f in _sample_fractions(animations, n)]
        self._ppt_last_t = 0.0
        self._ppt_sampling = True
        self._ppt_watch_cleanup()

    def update_to_time(self, t: float) -> None:
        super().update_to_time(t)  # type: ignore[misc]
        while self._ppt_times and t >= self._ppt_times[0] - 1e-9:
            self._ppt_times.pop(0)
            self._ppt_append(self._ppt_snapshot(), t - self._ppt_last_t)
            self._ppt_last_t = t

    def render(self, preview: bool = False) -> bool:
        result = super().render(preview)  # type: ignore[misc]
        if self.ppt_auto_export:
            self.ppt_export()
        return result

    # ------------------------------------------------------------ export

    def ppt_timeline(self) -> Timeline:
        """The captured keyframes, finalised for export."""
        tl = self._ppt_timeline
        final = self._ppt_snapshot()
        if not tl.keyframes:
            self._ppt_new_keyframe(final, 0, starts_play=False, marker=False)
        elif final.fingerprint() != self._ppt_last_fp:
            self._ppt_pending_start = True
            self._ppt_append(final, 0.0)
        self._ppt_last_fp = final.fingerprint()
        mode = self.ppt_mode
        if mode == "auto":
            mode = "markers" if self._ppt_markers_used else "click"
        tl.mode = mode  # type: ignore[assignment]
        tl.title = tl.title or type(self).__name__
        return tl

    def ppt_export(self, path: str | Path | None = None) -> Path:
        """Write the deck and return its path."""
        path = Path(
            path
            or self.ppt_output
            or Path(config.media_dir) / "pptx" / f"{type(self).__name__}.pptx"
        )
        tl = self.ppt_timeline()
        out, slides = write_deck(
            tl,
            path,
            template=self.ppt_template,
            background=self.ppt_background,
            native_shapes=self.ppt_native_shapes,
            group=self.ppt_group,
        )
        self.ppt_path, self.ppt_slide_count = out, slides
        logger.info(
            "manim-ppt: wrote %(n)d slides to %(path)s",
            {"n": slides, "path": str(out)},
        )
        return out

    # ---------------------------------------------------------- internals

    def _ppt_snapshot(self) -> Snapshot:
        return capture(self, self._ppt_namer)  # type: ignore[arg-type]

    def _ppt_drawn_keys(self) -> frozenset[str]:
        keys: set[str] = set()
        for anim in _leaf_animations(self.animations or []):  # type: ignore[attr-defined]
            if isinstance(anim, _DRAWING_ANIMATIONS):
                for mob in anim.mobject.get_family():
                    if mob.has_points():
                        keys.add(self._ppt_namer(mob))
        return frozenset(keys)

    def _ppt_watch_cleanup(self) -> None:
        """Snapshot the scene right before the first animation cleans up.

        At that instant every animation has reached alpha=1 but mobjects that
        are about to leave are still in the scene. That is the only moment to
        see where a ``FadeOut(shift=...)`` ends up, and it lets mobjects that
        ``ReplacementTransform`` & co. swap in take over the names of the
        ones they replace.
        """
        animations = self.animations  # type: ignore[attr-defined]
        if not animations:
            return
        anim = animations[0]
        original = anim.clean_up_from_scene

        def clean_up_from_scene(scene: Scene) -> Any:
            anim.clean_up_from_scene = original
            self._ppt_precleanup = self._ppt_snapshot()
            return original(scene)

        anim.clean_up_from_scene = clean_up_from_scene

    def _ppt_inherit_names(self, before: Snapshot, after: Snapshot) -> None:
        before_keys = {s.key for s in before.shapes}
        after_keys = {s.key for s in after.shapes}
        new = [s for s in after.shapes if s.key not in before_keys and s.visible]
        gone = [s for s in before.shapes if s.key not in after_keys and s.visible]
        if not new or not gone:
            return
        tol = 0.01 * config.frame_width
        pairs = []
        for i, n in enumerate(new):
            for j, g in enumerate(gone):
                d = bounds_distance(n, g)
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
        return any(mob.get_family_updaters() for mob in self.mobjects)  # type: ignore[attr-defined]

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
        tl.keyframes[-1].hold_ms += round(duration * 1000)

    def _ppt_append(self, snap: Snapshot, transition_s: float) -> None:
        tl = self._ppt_timeline
        ms = max(0, round(transition_s * 1000))
        if tl.keyframes and snap.fingerprint() == tl.keyframes[-1].snapshot.fingerprint():
            # No visible change: just extend the time spent on the last slide.
            tl.keyframes[-1].hold_ms += ms
            return
        self._ppt_new_keyframe(snap, ms, starts_play=self._ppt_pending_start, marker=False)
        self._ppt_pending_start = False

    def _ppt_new_keyframe(
        self, snap: Snapshot, transition_ms: int, *, starts_play: bool, marker: bool
    ) -> None:
        self._ppt_timeline.keyframes.append(
            Keyframe(
                snap,
                transition_ms=transition_ms,
                starts_play=starts_play,
                marker=marker,
                notes=self._ppt_notes,
                section=self._ppt_section,
                fade_keys=self._ppt_fade_keys,
            )
        )
        self._ppt_section = None


class PPTScene(PPTMixin, Scene):
    """Use exactly like ``Scene``; a ``.pptx`` is written next to the video.

    See :class:`PPTMixin` for the slide-flow rules and the ``ppt_*`` options.
    """
