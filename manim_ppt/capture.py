"""Read the current state of a manim scene into a renderer-independent
:class:`~manim_ppt.model.Snapshot`."""

from __future__ import annotations

import hashlib
import io
from typing import TYPE_CHECKING

import numpy as np
from manim import DR, UL, ImageMobject, VMobject
from manim.constants import CapStyleType, LineJointType

from .model import ImageShape, PathShape, Snapshot

if TYPE_CHECKING:
    from manim import Mobject, Scene

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


class MobjectNamer:
    """Gives every mobject a stable key for the whole scene.

    Strong references are kept so Python never recycles an ``id()`` for a
    different object during the render.
    """

    def __init__(self) -> None:
        self._names: dict[int, tuple[Mobject, str]] = {}
        self._by_key: dict[str, Mobject] = {}
        self._count = 0

    def __call__(self, mob: Mobject) -> str:
        entry = self._names.get(id(mob))
        if entry is None or entry[0] is not mob:
            entry = (mob, self._fresh())
            self._names[id(mob)] = entry
            self._by_key[entry[1]] = mob
        return entry[1]

    def _fresh(self) -> str:
        self._count += 1
        return f"m{self._count}"

    def transfer(self, key: str, new_mob: Mobject) -> None:
        """Give ``key`` to ``new_mob``; its previous owner gets a fresh key."""
        old = self._by_key.get(key)
        if old is not None and old is not new_mob:
            fresh = self._fresh()
            self._names[id(old)] = (old, fresh)
            self._by_key[fresh] = old
        self._names[id(new_mob)] = (new_mob, key)
        self._by_key[key] = new_mob

    def mobject(self, key: str) -> Mobject | None:
        return self._by_key.get(key)


_png_cache: dict[tuple[int, str], bytes] = {}


def _rgba(arr) -> tuple[float, float, float, float] | None:
    if arr is None or len(arr) == 0:
        return None
    r, g, b, a = (float(v) for v in arr[0])
    return (r, g, b, a)


def capture_vmobject(mob: VMobject, key: str) -> PathShape | None:
    subpaths = [sp[:, :2].astype(float) for sp in mob.get_subpaths() if len(sp) >= 4]
    if not subpaths:
        return None
    closed = [bool(np.allclose(sp[0], sp[-1], atol=1e-6)) for sp in subpaths]
    fill = _rgba(mob.get_fill_rgbas())
    stroke = _rgba(mob.get_stroke_rgbas())
    width = float(mob.get_stroke_width() or 0.0)
    if (fill is None or fill[3] <= 0.001) and (stroke is None or stroke[3] <= 0.001 or width <= 0):
        return None
    return PathShape(
        key=key,
        subpaths=subpaths,
        closed=closed,
        fill=fill,
        stroke=stroke,
        stroke_width=width,
        cap=_CAP.get(getattr(mob, "cap_style", CapStyleType.AUTO), "flat"),
        join=_JOIN.get(getattr(mob, "joint_type", LineJointType.AUTO), "miter"),
    )


def capture_image(mob: ImageMobject, key: str) -> ImageShape | None:
    from PIL import Image

    pixels = np.asarray(mob.get_pixel_array())
    digest = hashlib.md5(pixels.tobytes()).hexdigest()
    cache_key = (id(mob), digest)
    png = _png_cache.get(cache_key)
    if png is None:
        buf = io.BytesIO()
        mode = "RGBA" if pixels.ndim == 3 and pixels.shape[2] == 4 else None
        Image.fromarray(pixels.astype(np.uint8), mode=mode).save(buf, format="PNG")
        png = _png_cache[cache_key] = buf.getvalue()
    ul, dr = mob.get_corner(UL), mob.get_corner(DR)
    box = (float(ul[0]), float(dr[1]), float(dr[0]), float(ul[1]))
    if box[2] - box[0] <= 1e-6 or box[3] - box[1] <= 1e-6:
        return None
    return ImageShape(key=key, png=png, box=box)


def capture(scene: Scene, namer: MobjectNamer) -> Snapshot:
    shapes = []
    for mob in scene.get_mobject_family_members():
        shape = None
        if isinstance(mob, VMobject):
            if mob.has_points():
                shape = capture_vmobject(mob, namer(mob))
        elif isinstance(mob, ImageMobject):
            shape = capture_image(mob, namer(mob))
        if shape is not None:
            shapes.append(shape)
    bg = scene.camera.background_color
    try:
        rgb = tuple(float(v) for v in bg.to_rgb())
    except AttributeError:
        rgb = (0.0, 0.0, 0.0)
    return Snapshot(shapes=shapes, background=rgb)
