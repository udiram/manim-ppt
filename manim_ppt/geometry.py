"""Geometry that makes the exported shapes *native* rather than merely correct.

Two questions are answered here, both in slide space (x right, **y down**,
angles clockwise-positive, exactly like DrawingML's ``rot``):

* :func:`similarity_angle`: did a shape merely rotate between two slides? If so
  the rotation goes into the shape's transform (``<a:xfrm rot=...>``) instead
  of being baked into its points. PowerPoint then shows a real rotation handle,
  and Morph spins the shape along an arc instead of dragging every point in a
  straight line.
* :func:`classify`: is a path really a rectangle, rounded rectangle, ellipse,
  triangle or straight line? Then it is written as that preset shape, with
  the adjust handles and editing behaviour PowerPoint users expect.

Nothing here imports manim or python-pptx.
"""

from __future__ import annotations

import math
from dataclasses import dataclass

import numpy as np

REL_TOL = 2e-3  # tolerance relative to the size of the shape

_T = np.array([0.0, 0.25, 0.5, 0.75])
_BERNSTEIN = np.stack(
    [(1 - _T) ** 3, 3 * (1 - _T) ** 2 * _T, 3 * (1 - _T) * _T**2, _T**3], axis=1
)  # (4 samples, 4 control points)


@dataclass(frozen=True)
class Preset:
    kind: str  # "rect" | "roundRect" | "ellipse" | "triangle" | "line"
    adj: float | None = None  # the preset's adjust value as a 0..1 fraction

    def matches(self, other: Preset | None) -> bool:
        if other is None or other.kind != self.kind:
            return False
        if self.adj is None or other.adj is None:
            return self.adj is other.adj
        return abs(self.adj - other.adj) < 0.01


@dataclass
class LocalFrame:
    """A shape expressed the DrawingML way: an unrotated box plus a rotation
    about the center of that box."""

    theta: float
    x: float
    y: float
    w: float
    h: float
    subpaths: list[np.ndarray]  # control points relative to (x, y), unrotated


def rotation(theta: float) -> np.ndarray:
    c, s = math.cos(theta), math.sin(theta)
    return np.array([[c, -s], [s, c]])


def wrap(angle: float, period: float = 2 * math.pi) -> float:
    """Wrap ``angle`` into ``(-period/2, period/2]``."""
    a = math.fmod(angle, period)
    if a > period / 2:
        a -= period
    elif a <= -period / 2:
        a += period
    return a


def to_local(subpaths: list[np.ndarray], theta: float) -> LocalFrame:
    """Express ``subpaths`` in a box rotated by ``theta`` about its center."""
    if theta:
        inv = rotation(-theta).T
        local = [sp @ inv for sp in subpaths]
    else:
        local = [np.asarray(sp, dtype=float) for sp in subpaths]
    pts = np.concatenate(local)
    lo, hi = pts.min(axis=0), pts.max(axis=0)
    # DrawingML rotates about the box center, so the unrotated box has to sit
    # where the rotated center lands.
    center = rotation(theta) @ ((lo + hi) / 2) if theta else (lo + hi) / 2
    w, h = hi - lo
    return LocalFrame(
        theta=theta,
        x=float(center[0] - w / 2),
        y=float(center[1] - h / 2),
        w=float(w),
        h=float(h),
        subpaths=[sp - lo for sp in local],
    )


def similarity(
    src: list[np.ndarray], dst: list[np.ndarray], rel_tol: float = REL_TOL
) -> complex | None:
    """The rigid motion (plus uniform scale) taking ``src`` to ``dst`` point
    for point, as a complex factor: its argument is the rotation, its modulus
    the scale. ``None`` if the points were deformed instead."""
    if len(src) != len(dst) or any(a.shape != b.shape for a, b in zip(src, dst, strict=True)):
        return None
    p, q = np.concatenate(src), np.concatenate(dst)
    z = (p[:, 0] - p[:, 0].mean()) + 1j * (p[:, 1] - p[:, 1].mean())
    w = (q[:, 0] - q[:, 0].mean()) + 1j * (q[:, 1] - q[:, 1].mean())
    zz, ww = float(np.vdot(z, z).real), float(np.vdot(w, w).real)
    if zz <= 1e-12 or ww <= 1e-12:
        return None  # collapsed to a point: no orientation to speak of
    a = np.vdot(z, w) / zz  # least-squares complex scale: w ~ a * z
    residual = float(np.sum(np.abs(w - a * z) ** 2))
    if residual > rel_tol**2 * ww:
        return None
    return complex(a)


def similarity_angle(
    src: list[np.ndarray], dst: list[np.ndarray], rel_tol: float = REL_TOL
) -> float | None:
    """Rotation of the rigid motion taking ``src`` to ``dst``, or ``None``."""
    a = similarity(src, dst, rel_tol)
    return None if a is None else float(np.angle(a))


# --------------------------------------------------------------------------
# Path inspection
# --------------------------------------------------------------------------


def _curves(sp: np.ndarray) -> np.ndarray:
    return sp[: len(sp) // 4 * 4].reshape(-1, 4, 2)


def _samples(sp: np.ndarray) -> np.ndarray:
    """Points along the path, ending on its last anchor."""
    pts = np.einsum("st,ktd->ksd", _BERNSTEIN, _curves(sp)).reshape(-1, 2)
    return np.vstack([pts, sp[-1]])


def _straight_mask(sp: np.ndarray, tol: float) -> np.ndarray:
    """Which cubic segments of ``sp`` are straight lines?"""
    c = _curves(sp)
    chord = c[:, 3] - c[:, 0]
    length = np.linalg.norm(chord, axis=1)
    safe = np.where(length > 0, length, 1.0)
    out = np.ones(len(c), dtype=bool)
    for k in (1, 2):
        rel = c[:, k] - c[:, 0]
        cross = np.abs(rel[:, 0] * chord[:, 1] - rel[:, 1] * chord[:, 0]) / safe
        along = (rel * chord).sum(axis=1) / safe
        on_line = (cross <= tol) & (along >= -tol) & (along <= length + tol)
        at_point = np.linalg.norm(rel, axis=1) <= tol
        out &= np.where(length > tol, on_line, at_point)
    return out


def _vertices(sp: np.ndarray, closed: bool, tol: float) -> np.ndarray:
    """Corners of a path made of straight segments."""
    c = _curves(sp)
    pts = [c[0, 0], *c[:, 3]]
    out: list[np.ndarray] = []
    for p in pts:
        if not out or np.linalg.norm(p - out[-1]) > tol:
            out.append(p)
    if closed and len(out) > 1 and np.linalg.norm(out[0] - out[-1]) <= tol:
        out.pop()
    changed = True
    while changed and len(out) > 2:
        changed = False
        n = len(out)
        for i in range(n):
            if not closed and i in (0, n - 1):
                continue
            a, b, c2 = out[i - 1], out[i], out[(i + 1) % n]
            d = c2 - a
            length = np.linalg.norm(d)
            dist = (
                abs((b - a)[0] * d[1] - (b - a)[1] * d[0]) / length
                if length > tol
                else np.linalg.norm(b - a)
            )
            if dist <= tol:
                del out[i]
                changed = True
                break
    return np.array(out)


def _size(subpaths: list[np.ndarray]) -> float:
    pts = np.concatenate(subpaths)
    return float(np.linalg.norm(pts.max(axis=0) - pts.min(axis=0)))


def is_straight_line(subpaths: list[np.ndarray], closed: list[bool]) -> bool:
    """True for a single open path whose points all lie on one segment."""
    if len(subpaths) != 1 or closed[0]:
        return False
    tol = REL_TOL * _size(subpaths)
    if tol <= 0:
        return False
    sp = subpaths[0]
    return bool(_straight_mask(sp, tol).all()) and len(_vertices(sp, False, tol)) == 2


# --------------------------------------------------------------------------
# Preset recognition
# --------------------------------------------------------------------------


def _winds_once(pts: np.ndarray) -> bool:
    ang = np.unwrap(np.arctan2(pts[:, 1], pts[:, 0]))
    return abs(abs(ang[-1] - ang[0]) - 2 * math.pi) < 0.3


def _rounded_box_distance(p: np.ndarray, half: np.ndarray, r: float) -> np.ndarray:
    q = np.abs(p) - (half - r)
    outside = np.linalg.norm(np.maximum(q, 0.0), axis=1)
    inside = np.minimum(np.maximum(q[:, 0], q[:, 1]), 0.0)
    return outside + inside - r


def classify(frame: LocalFrame, closed: list[bool]) -> Preset | None:
    """Name the preset shape that ``frame`` draws, if there is one."""
    if len(frame.subpaths) != 1:
        return None
    sp, w, h = frame.subpaths[0], frame.w, frame.h
    size = math.hypot(w, h)
    tol = REL_TOL * size
    if tol <= 0 or len(sp) < 4:
        return None

    if not closed[0]:
        flat = h <= tol and w > tol and abs(sp[0, 0] - sp[-1, 0]) >= w - tol
        return Preset("line") if flat and _straight_mask(sp, tol).all() else None

    if min(w, h) <= tol:
        return None
    half = np.array([w / 2, h / 2])
    pts = _samples(sp) - half
    if not _winds_once(pts):
        return None

    # Ellipse: every sample satisfies (x/a)^2 + (y/b)^2 = 1.
    q = (pts[:, 0] / half[0]) ** 2 + (pts[:, 1] / half[1]) ** 2
    if np.abs(q - 1).max() < 0.012:
        return Preset("ellipse")

    straight = _straight_mask(sp, tol)

    # (Rounded) rectangle: the straight run along the top edge stops one
    # corner radius short of each side.
    c = _curves(sp)[straight]
    on_top = c[(c[:, 0, 1] <= tol) & (c[:, 3, 1] <= tol)]
    short = min(w, h)
    if len(on_top):
        xs = on_top[:, (0, 3), 0]
        r = max(float(xs.min()), float(w - xs.max()), 0.0)
    else:
        r = short / 2  # fully rounded ends
    if r <= short / 2 + tol:
        r = min(r, short / 2)
        if np.abs(_rounded_box_distance(pts, half, r)).max() <= tol:
            return Preset("rect") if r <= 2 * tol else Preset("roundRect", r / short)

    # Triangle with a flat base at the bottom and its apex on the top edge.
    if straight.all():
        v = _vertices(sp, True, tol)
        if len(v) == 3:
            base = v[v[:, 1] >= h - tol]
            apex = v[v[:, 1] <= tol]
            if (
                len(base) == 2
                and len(apex) == 1
                and base[:, 0].min() <= tol
                and base[:, 0].max() >= w - tol
            ):
                return Preset("triangle", float(np.clip(apex[0, 0] / w, 0.0, 1.0)))
    return None


def _triangle_angle(v: np.ndarray) -> float:
    """Pick the base edge that makes the triangle a PowerPoint triangle
    (apex up): isosceles if possible, otherwise the least rotated."""
    best: tuple[int, float, float] | None = None
    for i in range(3):
        a, b, apex = v[i], v[(i + 1) % 3], v[(i + 2) % 3]
        d = b - a
        theta = math.atan2(d[1], d[0])
        local = rotation(-theta) @ (apex - a)
        if local[1] > 0:  # apex below the base (y is down): flip the frame
            theta = wrap(theta + math.pi)
            local = rotation(-theta) @ (apex - b)
        length = float(np.linalg.norm(d))
        adj = local[0] / length if length else 0.5
        key = (0 if abs(adj - 0.5) < 0.01 else 1, abs(theta), theta)
        if best is None or key < best:
            best = key
    return best[2] if best else 0.0


def natural_angle(subpaths: list[np.ndarray], closed: list[bool]) -> float:
    """The rotation at which a shape is a native preset (0 if it never is).

    ``Square().rotate(PI / 6)`` should come out as a rectangle rotated by 30
    degrees, not as a diamond-ish freeform.
    """
    if len(subpaths) != 1:
        return 0.0
    sp = subpaths[0]
    tol = REL_TOL * _size(subpaths)
    if tol <= 0 or len(sp) < 4:
        return 0.0

    candidates = [0.0]
    straight = _straight_mask(sp, tol)
    if straight.all():
        v = _vertices(sp, closed[0], tol)
        if len(v) == 2 and not closed[0]:
            d = v[1] - v[0]
            return wrap(math.atan2(d[1], d[0]), math.pi)
        if len(v) == 3 and closed[0]:
            candidates.append(_triangle_angle(v))
        elif len(v) == 4 and closed[0]:
            d = v[1] - v[0]
            candidates.append(wrap(math.atan2(d[1], d[0]), math.pi / 2))
    elif closed[0]:
        c = _curves(sp)[straight]
        if len(c):
            d = c[:, 3] - c[:, 0]
            longest = d[np.argmax(np.linalg.norm(d, axis=1))]
            candidates.append(wrap(math.atan2(longest[1], longest[0]), math.pi / 2))
        pts = _samples(sp)
        cov = np.cov((pts - pts.mean(axis=0)).T)
        vals, vecs = np.linalg.eigh(cov)
        if vals[1] > 0 and (vals[1] - vals[0]) / vals[1] > 1e-3:  # not a circle
            major = vecs[:, 1]
            candidates.append(wrap(math.atan2(major[1], major[0]), math.pi))

    for theta in candidates:
        if classify(to_local(subpaths, theta), closed) is not None:
            return theta
    return 0.0
