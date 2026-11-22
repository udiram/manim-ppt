"""Hand-built shapes for tests that do not need manim."""

from itertools import pairwise

import numpy as np

from manim_ppt.model import PathShape

FW, FH = 14.222222222222221, 8.0
SLIDE_W, SLIDE_H = 12_192_000, 6_858_000


def polygon(key, corners, fill=(1, 0, 0, 1), closed=True, **kwargs):
    """A PathShape with straight edges through ``corners`` (manim-style cubics)."""
    corners = [np.array(c, dtype=float) for c in corners]
    pairs = list(pairwise([*corners, corners[0]] if closed else corners))
    pts = []
    for a, b in pairs:
        pts += [a, a + (b - a) / 3, a + 2 * (b - a) / 3, b]
    kwargs.setdefault("stroke", (1, 1, 1, 1))
    kwargs.setdefault("stroke_width", 4)
    return PathShape(key=key, subpaths=[np.array(pts)], closed=[closed], fill=fill, **kwargs)


def square(key, cx=0.0, cy=0.0, r=1.0, angle=0.0, **kwargs):
    c, s = np.cos(angle), np.sin(angle)
    corners = [(-r, -r), (r, -r), (r, r), (-r, r)]
    return polygon(key, [(cx + c * x - s * y, cy + s * x + c * y) for x, y in corners], **kwargs)


def circle(key, cx=0.0, cy=0.0, r=1.0, rx=None, n=8, **kwargs):
    rx = r if rx is None else rx
    k = 4 / 3 * np.tan(2 * np.pi / n / 4)
    pts = []
    for i in range(n):
        a0, a1 = 2 * np.pi * i / n, 2 * np.pi * (i + 1) / n
        p0, p1 = np.array([np.cos(a0), np.sin(a0)]), np.array([np.cos(a1), np.sin(a1)])
        t0, t1 = np.array([-np.sin(a0), np.cos(a0)]), np.array([-np.sin(a1), np.cos(a1)])
        pts += [p0, p0 + k * t0, p1 - k * t1, p1]
    arr = np.array(pts) * [rx, r] + [cx, cy]
    kwargs.setdefault("stroke", (1, 1, 1, 1))
    kwargs.setdefault("stroke_width", 4)
    kwargs.setdefault("fill", (0, 0, 1, 1))
    return PathShape(key=key, subpaths=[arr], closed=[True], **kwargs)


def blob(key, cx=0.0, cy=0.0, scale=1.0, **kwargs):
    """An irregular closed outline that is no preset shape."""
    return polygon(
        key,
        [
            (cx + scale * x, cy + scale * y)
            for x, y in [(0, 0), (2, 0.2), (1.2, 0.9), (2.1, 2), (0.3, 1.4)]
        ],
        **kwargs,
    )
