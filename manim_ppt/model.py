"""Renderer-independent data model for captured keyframes.

Nothing in this module imports manim, so the PowerPoint writer and the
timeline logic can be unit-tested (and reused) without a manim install.
All coordinates are in manim scene units (x right, y up, origin at the
frame center).
"""

from __future__ import annotations

import hashlib
from dataclasses import dataclass, field

import numpy as np

RGBA = tuple[float, float, float, float]


@dataclass
class PathShape:
    """One manim VMobject (a single family member that has points)."""

    key: str
    # Each subpath is an (4k, 2) array of cubic Bezier control points laid out
    # the manim way: anchor, handle, handle, anchor, anchor, handle, ...
    subpaths: list[np.ndarray]
    closed: list[bool]
    fill: RGBA | None
    stroke: RGBA | None
    stroke_width: float  # in scene units
    cap: str = "rnd"  # DrawingML: rnd | sq | flat
    join: str = "round"  # round | miter | bevel

    def bounds(self) -> tuple[float, float, float, float]:
        pts = np.concatenate(self.subpaths, axis=0)
        x0, y0 = pts.min(axis=0)
        x1, y1 = pts.max(axis=0)
        return float(x0), float(y0), float(x1), float(y1)

    def fingerprint(self, h: "hashlib._Hash") -> None:
        h.update(self.key.encode())
        for sp in self.subpaths:
            h.update(np.round(sp, 5).tobytes())
        h.update(repr((self.fill, self.stroke, round(self.stroke_width, 5))).encode())


@dataclass
class ImageShape:
    """A raster ImageMobject, placed by its axis-aligned bounding box."""

    key: str
    png: bytes
    box: tuple[float, float, float, float]  # x0, y0, x1, y1 (scene units)

    def bounds(self) -> tuple[float, float, float, float]:
        return self.box

    def fingerprint(self, h: "hashlib._Hash") -> None:
        h.update(self.key.encode())
        h.update(hashlib.md5(self.png).digest())
        h.update(np.round(np.array(self.box), 5).tobytes())


Shape = PathShape | ImageShape


@dataclass
class Snapshot:
    """Everything visible on screen at one instant, in paint order."""

    shapes: list[Shape]
    background: tuple[float, float, float] = (0.0, 0.0, 0.0)

    def fingerprint(self) -> str:
        h = hashlib.sha1()
        for s in self.shapes:
            s.fingerprint(h)
        h.update(repr(self.background).encode())
        return h.hexdigest()


@dataclass
class Keyframe:
    """One slide.

    ``transition_ms`` is the Morph duration *into* this slide (0 means a hard
    cut). The flags describe how the slide is *left*.
    """

    snapshot: Snapshot
    transition_ms: int = 0
    hold_ms: int = 0  # extra time to sit on this slide before auto-advancing
    starts_play: bool = False  # the transition into this slide begins a play()
    marker: bool = False  # next_slide() was called while this slide was last
    notes: str | None = None


@dataclass
class Timeline:
    keyframes: list[Keyframe] = field(default_factory=list)
    frame_width: float = 14.222222222222221
    frame_height: float = 8.0
    # "click": every play() waits for a click; holds are ignored.
    # "markers": only next_slide() waits for a click; waits become timed holds.
    mode: str = "click"
