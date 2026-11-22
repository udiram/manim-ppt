"""Read the current state of a manim scene into a renderer-independent
:class:`~manim_ppt.model.Snapshot`."""

from __future__ import annotations

import hashlib
import io
import math
from collections import Counter
from typing import TYPE_CHECKING, Any

import numpy as np
from manim import Group, ImageMobject, Mobject, VDict, VGroup, VMobject, config
from manim.constants import CapStyleType, LineJointType

from .model import RGBA, Gradient, ImageShape, PathShape, Shape, Snapshot

if TYPE_CHECKING:
    from collections.abc import Iterable

    from manim import Scene

_CAP = {
    CapStyleType.AUTO: "flat",
    CapStyleType.ROUND: "rnd",
    CapStyleType.BUTT: "flat",
    CapStyleType.SQUARE: "sq",
}
_JOIN = {
    LineJointType.AUTO: "miter",
    LineJointType.ROUND: "round",
    LineJointType.BEVEL: "bevel",
    LineJointType.MITER: "miter",
}

# Containers that say nothing about what they hold. Their members are named
# after the nearest meaningful ancestor instead ("Axes 1.4", "Text 2.7").
_GENERIC: frozenset[type] = frozenset({Mobject, VMobject, Group, VGroup, VDict})


class MobjectNamer:
    """Gives every mobject a stable, readable name for the whole scene.

    The names end up in PowerPoint's Selection Pane, and they are what Morph
    uses to pair a shape on one slide with itself on the next, so they must
    never change while a mobject lives.

    * A mobject on its own is named after its class: ``Circle 1``, ``Circle 2``.
    * Parts of a compound mobject are numbered under it: ``Text 1.1``,
      ``Text 1.2``... The number is the part's position, which keeps names
      steady even when the parts are rebuilt from scratch on every frame
      (``always_redraw``, ``DecimalNumber.set_value``, ``become``).

    Strong references are kept so Python never recycles an ``id()`` for a
    different object during the render.
    """

    def __init__(self) -> None:
        self._labels: dict[int, tuple[Mobject, str]] = {}
        self._owner: dict[str, Mobject] = {}
        self._class_count: Counter[str] = Counter()
        self._child_count: Counter[str] = Counter()
        self._group: dict[int, str] = {}

    def __call__(self, mob: Mobject) -> str:
        return self._get(mob) or self._fresh(mob)

    def mobject(self, key: str) -> Mobject | None:
        return self._owner.get(key)

    def group(self, mob: Mobject) -> str | None:
        """Name of the compound mobject ``mob`` was last seen as a part of."""
        return self._group.get(id(mob))

    def transfer(self, key: str, new_mob: Mobject) -> None:
        """Give ``key`` to ``new_mob``; its previous owner gets a fresh key."""
        old = self._owner.get(key)
        if old is new_mob:
            return
        self._forget(new_mob)
        if old is not None:
            del self._labels[id(old)]
            del self._owner[key]
            self._fresh(old)
        self._set(new_mob, key)

    def assign(self, roots: Iterable[Mobject]) -> None:
        """Name every mobject reachable from ``roots`` that has points."""
        live: set[int] = set()
        parts: dict[str, list[Mobject]] = {}
        self._group = {}

        def walk(mob: Mobject, context: str | None) -> None:
            if id(mob) in live:
                return
            live.add(id(mob))
            if context is None and (type(mob) not in _GENERIC or mob.has_points()):
                label = self(mob)
                self._group[id(mob)] = label
                if type(mob) not in _GENERIC:
                    context = label
            elif context is not None and mob.has_points():
                parts.setdefault(context, []).append(mob)
                self._group[id(mob)] = context
            for sub in mob.submobjects:
                walk(sub, context)

        for root in roots:
            walk(root, None)

        for context, members in parts.items():
            for position, mob in enumerate(members, start=1):
                if self._get(mob) is not None:
                    continue
                label = f"{context}.{position}"
                owner = self._owner.get(label)
                if owner is not None and id(owner) in live:
                    # The slot belongs to a part that is still around.
                    self._child_count[context] = max(self._child_count[context], len(members)) + 1
                    label = f"{context}.{self._child_count[context]}"
                elif owner is not None:
                    self._forget(owner)  # rebuilt in place: take over its name
                self._child_count[context] = max(self._child_count[context], position)
                self._set(mob, label)

    def _get(self, mob: Mobject) -> str | None:
        entry = self._labels.get(id(mob))
        return entry[1] if entry is not None and entry[0] is mob else None

    def _set(self, mob: Mobject, label: str) -> None:
        self._labels[id(mob)] = (mob, label)
        self._owner[label] = mob

    def _forget(self, mob: Mobject) -> None:
        label = self._get(mob)
        if label is not None:
            del self._labels[id(mob)]
            if self._owner.get(label) is mob:
                del self._owner[label]

    def _fresh(self, mob: Mobject) -> str:
        name = type(mob).__name__
        while True:
            self._class_count[name] += 1
            label = f"{name} {self._class_count[name]}"
            if label not in self._owner:
                self._set(mob, label)
                return label


class _View:
    """Where the camera is looking, relative to the default frame.

    Lets ``MovingCameraScene`` zooms and pans come through: everything is
    captured the way the camera sees it.
    """

    def __init__(self, camera: Any) -> None:
        center = np.asarray(getattr(camera, "frame_center", (0.0, 0.0, 0.0)), dtype=float)
        self.center = center[:2]
        width = float(getattr(camera, "frame_width", config.frame_width) or config.frame_width)
        height = float(getattr(camera, "frame_height", config.frame_height) or config.frame_height)
        self.scale = np.array([config.frame_width / width, config.frame_height / height])
        self.zoom = float(self.scale[0])
        # A 3D camera only refreshes its orientation when it draws a frame;
        # make sure a capture between frames sees the current one.
        refresh = getattr(camera, "reset_rotation_matrix", None)
        if refresh is not None:
            refresh()
        self._project = getattr(camera, "transform_points_pre_display", None)
        self._fill = getattr(camera, "get_fill_rgbas", None)
        self._stroke = getattr(camera, "get_stroke_rgbas", None)

    def transform(self, mob: Mobject, points: np.ndarray) -> np.ndarray:
        """What the camera does to points before drawing (3D projection)."""
        return self._project(mob, points) if self._project is not None else points

    def flatten(self, points: np.ndarray) -> np.ndarray:
        """Scene points -> 2D points as they appear in the default frame."""
        return (np.asarray(points, dtype=float)[:, :2] - self.center) * self.scale

    def fill_rgbas(self, mob: VMobject) -> np.ndarray:
        return self._fill(mob) if self._fill is not None else mob.get_fill_rgbas()

    def stroke_rgbas(self, mob: VMobject) -> np.ndarray:
        return self._stroke(mob) if self._stroke is not None else mob.get_stroke_rgbas()


_png_cache: dict[tuple[int, str], bytes] = {}
_MIN_IMAGE_SIDE = 512  # pixels


def _paint(rgbas: Any, mob: VMobject, view: _View) -> tuple[RGBA | None, Gradient | None]:
    """Split manim's color array into a solid color and an optional gradient."""
    if rgbas is None or len(rgbas) == 0:
        return None, None
    stops = [tuple(float(v) for v in row[:4]) for row in rgbas]
    first: RGBA = stops[0]  # type: ignore[assignment]
    if len(stops) == 1 or all(np.allclose(s, first, atol=1e-4) for s in stops):
        return first, None
    start, end = mob.get_gradient_start_and_end_points()
    direction = (np.asarray(end, dtype=float) - np.asarray(start, dtype=float))[:2] * view.scale
    if float(np.linalg.norm(direction)) < 1e-9:
        return first, None
    gradient = Gradient(tuple(stops), (float(direction[0]), float(direction[1])))  # type: ignore[arg-type]
    return first, gradient


def capture_vmobject(
    mob: VMobject, key: str, view: _View, group: str | None = None
) -> PathShape | None:
    points = view.transform(mob, mob.points)
    subpaths = [
        view.flatten(sp[: len(sp) // 4 * 4])
        for sp in mob.gen_subpaths_from_points_2d(points)
        if len(sp) >= 4
    ]
    if not subpaths:
        return None
    closed = [bool(np.allclose(sp[0], sp[-1], atol=1e-6)) for sp in subpaths]
    fill, fill_gradient = _paint(view.fill_rgbas(mob), mob, view)
    stroke, stroke_gradient = _paint(view.stroke_rgbas(mob), mob, view)
    return PathShape(
        key=key,
        subpaths=subpaths,
        closed=closed,
        fill=fill,
        stroke=stroke,
        stroke_width=float(mob.get_stroke_width() or 0.0) * view.zoom,
        cap=_CAP.get(getattr(mob, "cap_style", CapStyleType.AUTO), "flat"),
        join=_JOIN.get(getattr(mob, "joint_type", LineJointType.AUTO), "miter"),
        fill_gradient=fill_gradient,
        stroke_gradient=stroke_gradient,
        group=group,
    )


def capture_image(mob: ImageMobject, key: str, view: _View) -> ImageShape | None:
    from PIL import Image

    if len(mob.points) < 3:
        return None
    pixels = np.asarray(mob.get_pixel_array())
    digest = hashlib.md5(pixels.tobytes(), usedforsecurity=False).hexdigest()
    cache_key = (id(mob), digest)
    png = _png_cache.get(cache_key)
    if png is None:
        image = Image.fromarray(pixels.astype(np.uint8))
        # PowerPoint smooths pictures when it enlarges them. Pre-scale small
        # ones the way manim would draw them (e.g. crisp pixels for "nearest").
        factor = math.ceil(_MIN_IMAGE_SIDE / max(1, min(image.size)))
        if factor > 1:
            resample = getattr(mob, "resampling_algorithm", Image.Resampling.BICUBIC)
            image = image.resize((image.width * factor, image.height * factor), resample)
        buf = io.BytesIO()
        image.save(buf, format="PNG")
        png = _png_cache[cache_key] = buf.getvalue()
    ul, ur, dl = view.flatten(view.transform(mob, mob.points[:3]))
    across, down = ur - ul, dl - ul
    width, height = float(np.linalg.norm(across)), float(np.linalg.norm(down))
    if width <= 1e-6 or height <= 1e-6:
        return None
    center = (ur + dl) / 2
    return ImageShape(
        key=key,
        png=png,
        center=(float(center[0]), float(center[1])),
        width=width,
        height=height,
        angle=math.atan2(across[1], across[0]),
    )


def _display_list(scene: Scene) -> list[Mobject]:
    """Every mobject with points, in the order the camera paints them."""
    display = getattr(scene.camera, "get_mobjects_to_display", None)
    if display is not None:
        try:
            return list(display(scene.mobjects))
        except Exception:  # pragma: no cover - exotic cameras
            pass
    return list(scene.get_mobject_family_members())


def capture(scene: Scene, namer: MobjectNamer) -> Snapshot:
    view = _View(scene.camera)
    namer.assign(scene.mobjects)
    shapes: list[Shape] = []
    for mob in _display_list(scene):
        shape: Shape | None = None
        if isinstance(mob, VMobject):
            if mob.has_points():
                shape = capture_vmobject(mob, namer(mob), view, namer.group(mob))
        elif isinstance(mob, ImageMobject):
            shape = capture_image(mob, namer(mob), view)
        if shape is not None:
            shapes.append(shape)
    bg = scene.camera.background_color
    try:
        rgb = tuple(float(v) for v in bg.to_rgb())
    except AttributeError:
        rgb = (0.0, 0.0, 0.0)
    return Snapshot(shapes=shapes, background=rgb)  # type: ignore[arg-type]
