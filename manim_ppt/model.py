"""Renderer-independent data model for captured keyframes.

Nothing in this module imports manim, so the PowerPoint writer and the
timeline logic can be unit-tested (and reused) without a manim install.
All coordinates are in manim scene units (x right, y up, origin at the
frame center).
"""

from __future__ import annotations

import hashlib
import math
from dataclasses import dataclass, field, replace
from typing import TYPE_CHECKING, Literal

import numpy as np

if TYPE_CHECKING:
    from hashlib import _Hash

RGBA = tuple[float, float, float, float]
Bounds = tuple[float, float, float, float]  # x0, y0, x1, y1

_VISIBLE_ALPHA = 1e-3
# Anything smaller than this (in scene units) has collapsed to a point: the
# start of a GrowFromCenter, the end of an Uncreate...
COLLAPSE_SIZE = 0.01


@dataclass(frozen=True)
class Gradient:
    """A linear gradient: evenly spaced ``stops`` along ``direction``."""

    stops: tuple[RGBA, ...]
    direction: tuple[float, float]  # scene-space vector, first stop -> last stop

    def faded(self) -> Gradient:
        return Gradient(tuple((r, g, b, 0.0) for r, g, b, _ in self.stops), self.direction)

    @property
    def alpha(self) -> float:
        return max(s[3] for s in self.stops)


@dataclass
class PathShape:
    """One manim VMobject (a single family member that has points)."""

    key: str
    # Each subpath is a (4k, 2) array of cubic Bezier control points laid out
    # the manim way: anchor, handle, handle, anchor, anchor, handle, ...
    subpaths: list[np.ndarray]
    closed: list[bool]
    fill: RGBA | None
    stroke: RGBA | None
    stroke_width: float  # manim stroke units; 1 unit = 0.01 scene units
    cap: str = "rnd"  # DrawingML: rnd | sq | flat
    join: str = "round"  # round | miter | bevel
    fill_gradient: Gradient | None = None
    stroke_gradient: Gradient | None = None
    # Name of the compound mobject this shape is a part of ("Text 1" for the
    # glyph "Text 1.4"); a mobject on its own is its own group.
    group: str | None = None

    def bounds(self) -> Bounds:
        pts = np.concatenate(self.subpaths, axis=0)
        x0, y0 = pts.min(axis=0)
        x1, y1 = pts.max(axis=0)
        return float(x0), float(y0), float(x1), float(y1)

    @property
    def has_fill(self) -> bool:
        alpha = self.fill_gradient.alpha if self.fill_gradient else (self.fill or (0, 0, 0, 0))[3]
        return alpha > _VISIBLE_ALPHA

    @property
    def has_stroke(self) -> bool:
        alpha = (
            self.stroke_gradient.alpha if self.stroke_gradient else (self.stroke or (0, 0, 0, 0))[3]
        )
        return alpha > _VISIBLE_ALPHA and self.stroke_width > 0

    @property
    def collapsed(self) -> bool:
        x0, y0, x1, y1 = self.bounds()
        return max(x1 - x0, y1 - y0) < COLLAPSE_SIZE

    @property
    def visible(self) -> bool:
        return (self.has_fill or self.has_stroke) and not self.collapsed

    def faded(self) -> PathShape:
        """A fully transparent copy: same geometry and colors, zero opacity.

        Morph tweens opacity, so a faded shape on one slide and its visible
        twin on the next read as a fade (combined with whatever move, scale
        or rotation separates the two).
        """

        def clear(c: RGBA | None) -> RGBA | None:
            return None if c is None else (c[0], c[1], c[2], 0.0)

        return replace(
            self,
            fill=clear(self.fill),
            stroke=clear(self.stroke),
            fill_gradient=self.fill_gradient.faded() if self.fill_gradient else None,
            stroke_gradient=self.stroke_gradient.faded() if self.stroke_gradient else None,
        )

    def fingerprint(self, h: _Hash) -> None:
        h.update(self.key.encode())
        for sp in self.subpaths:
            h.update(np.round(sp, 5).tobytes())
        style = (
            self.fill,
            self.stroke,
            round(self.stroke_width, 5),
            self.fill_gradient,
            self.stroke_gradient,
        )
        h.update(repr(style).encode())


@dataclass
class ImageShape:
    """A raster ImageMobject: a picture with a center, a size and a rotation."""

    key: str
    png: bytes
    center: tuple[float, float]
    width: float
    height: float
    angle: float = 0.0  # radians, counter-clockwise (manim convention)

    visible = True

    def bounds(self) -> Bounds:
        c, s = abs(math.cos(self.angle)), abs(math.sin(self.angle))
        hw = (self.width * c + self.height * s) / 2
        hh = (self.width * s + self.height * c) / 2
        return (
            self.center[0] - hw,
            self.center[1] - hh,
            self.center[0] + hw,
            self.center[1] + hh,
        )

    def fingerprint(self, h: _Hash) -> None:
        h.update(self.key.encode())
        h.update(hashlib.md5(self.png, usedforsecurity=False).digest())
        placement = np.array([*self.center, self.width, self.height, self.angle])
        h.update(np.round(placement, 5).tobytes())


Shape = PathShape | ImageShape


@dataclass
class Snapshot:
    """Everything on screen at one instant, in paint order.

    Fully transparent shapes are included: they matter when the same mobject
    is visible on a neighbouring slide (see :meth:`PathShape.faded`).
    """

    shapes: list[Shape]
    background: tuple[float, float, float] = (0.0, 0.0, 0.0)

    def fingerprint(self) -> str:
        """Hash of what an audience would see (invisible shapes excluded)."""
        h = hashlib.sha1(usedforsecurity=False)
        for s in self.shapes:
            if s.visible:
                s.fingerprint(h)
        h.update(repr(self.background).encode())
        return h.hexdigest()


@dataclass
class Keyframe:
    """One moment the deck stops on (usually one slide).

    ``transition_ms`` is the Morph duration *into* this slide (0 means a hard
    cut). The flags describe how the slide is *left*.

    ``snapshot`` is the scene as the previous animation left it. ``departure``
    is the same picture as the next animation starts from it: it looks
    identical, but carries that animation's starting geometry (a point for
    ``GrowFromCenter``, re-aligned points for ``Transform``...), which gives
    Morph something to tween from.
    """

    snapshot: Snapshot
    departure: Snapshot | None = None
    transition_ms: int = 0
    hold_ms: int = 0  # extra time to sit on this slide before auto-advancing
    starts_play: bool = False  # the transition into this slide begins a play()
    marker: bool = False  # next_slide() was called while this slide was last
    notes: str | None = None
    section: str | None = None  # a PowerPoint section starts at this slide
    # Mobjects that are being drawn in/out here (Create, Write...). When such
    # a shape has collapsed to a point it is left off the slide, so that Morph
    # fades it instead of zooming it out of a dot.
    fade_keys: frozenset[str] = frozenset()


Mode = Literal["click", "markers"]


@dataclass
class Timeline:
    keyframes: list[Keyframe] = field(default_factory=list)
    frame_width: float = 14.222222222222221
    frame_height: float = 8.0
    # "click": every play() waits for a click; holds are ignored.
    # "markers": only next_slide() waits for a click; waits become timed holds.
    mode: Mode = "click"
    title: str | None = None


def bounds_distance(a: Shape, b: Shape) -> float | None:
    """Largest edge-to-edge gap between two bounding boxes (``None`` if the
    shapes are of different kinds and so can never stand in for each other)."""
    if type(a) is not type(b):
        return None
    return max(abs(p - q) for p, q in zip(a.bounds(), b.bounds(), strict=True))
