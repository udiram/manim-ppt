"""Turn a :class:`~manim_ppt.model.Timeline` into a .pptx built from native
shapes, chained together with PowerPoint's Morph transition.

How it works
------------
* Every keyframe becomes a slide, and every manim VMobject a PowerPoint shape
  that keeps one name for the whole deck. Morph pairs objects whose names
  start with ``!!``, and tweens position, size, rotation, geometry, color and
  opacity between slides instead of cross-fading.
* Shapes are as native as their geometry allows (see :mod:`.geometry`):
  rectangles, rounded rectangles, ellipses, triangles and lines are written
  as PowerPoint presets, and rotation is stored as a real rotation. Everything
  else is a freeform (``a:custGeom``) built from ``a:cubicBezTo`` segments,
  which are exactly manim's cubic Beziers.
* A compound mobject (a word, an equation, a set of axes) that moves as a
  whole is written as one group, which is both what a PowerPoint user expects
  to click on and what Morph animates most cleanly.
* The slide-advance flags reproduce the scene's timing: sub-steps of one
  ``play()`` advance automatically, and a click is needed between plays (or
  only at ``next_slide()`` markers).

Which form a shape takes is decided per transition, because Morph is picky:
it scales presets and groups flawlessly but mangles complex freeforms as they
change size, and it can only bend one outline into another if both are
freeforms with matching points. When the form a shape needs for the
transition *into* a slide differs from the one it needs on the way *out*, the
slide is written twice with an instant, invisible cut in between.
"""

from __future__ import annotations

import datetime as dt
import io
import math
import uuid
from collections import Counter
from collections.abc import Sequence
from dataclasses import dataclass, replace
from itertools import pairwise
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape

import numpy as np
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Emu

from .geometry import (
    LocalFrame,
    Preset,
    classify,
    is_straight_line,
    natural_angle,
    rotation,
    similarity,
    similarity_angle,
    to_local,
)
from .model import (
    Gradient,
    ImageShape,
    Keyframe,
    PathShape,
    Shape,
    Snapshot,
    Timeline,
    bounds_distance,
)

SLIDE_WIDTH_EMU = 12_192_000  # 13.333 in, the PowerPoint 16:9 default
MANIM_STROKE_UNIT = 0.01  # cairo: line width = stroke_width * 0.01 scene units
SEED_SCALE = 0.02  # size of the stand-in for a shape that collapsed to a point

NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
NS_P159 = "http://schemas.microsoft.com/office/powerpoint/2015/09/main"
SECTION_EXT_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"
MORPH_PREFIX = "!!"

_NS = nsdecls("a", "p")


@dataclass
class State:
    """One static picture: a keyframe as the deck arrives at it, or as it
    leaves it (see :attr:`Keyframe.departure`)."""

    keyframe: int
    snapshot: Snapshot
    fade_keys: frozenset[str] = frozenset()


def expand(keyframes: list[Keyframe]) -> tuple[list[State], list[bool]]:
    """Unroll keyframes into states, and say which neighbours Morph connects
    (the others are separated by a hard cut)."""
    states: list[State] = []
    for k, kf in enumerate(keyframes):
        states.append(State(k, kf.snapshot, kf.fade_keys))
        if kf.departure is not None:
            states.append(State(k, kf.departure, kf.fade_keys))
    morph = [
        a.keyframe != b.keyframe and keyframes[b.keyframe].transition_ms > 0
        for a, b in pairwise(states)
    ]
    return states, morph


# --------------------------------------------------------------------------
# Shape naming across slides
# --------------------------------------------------------------------------


def resolve_names(keyframes: Sequence[Keyframe | State], frame_width: float) -> list[list[str]]:
    """Return, for each keyframe, the Morph name of each of its shapes.

    Names normally come straight from the mobject identity. Animations such as
    ``ReplacementTransform`` or ``TransformMatchingTex`` swap one mobject for
    another at the very end, which would make Morph cross-fade. To avoid that,
    a shape that newly appears inherits the name of a shape that just
    disappeared, provided the two are (nearly) the same geometry.
    """
    tol = 0.01 * frame_width
    alias: dict[str, str] = {}
    out: list[list[str]] = []
    prev: list[tuple[str, Shape]] = []

    for kf in keyframes:
        shapes = kf.snapshot.shapes
        names = [alias.get(s.key, s.key) for s in shapes]
        # Guard against two shapes claiming one name on a slide.
        seen: set[str] = set()
        for i, s in enumerate(shapes):
            if names[i] in seen:
                names[i] = s.key
            seen.add(names[i])

        prev_names = {n for n, _ in prev}
        cur_names = set(names)
        new_idx = [i for i, n in enumerate(names) if n not in prev_names]
        gone = [(n, s) for n, s in prev if n not in cur_names]

        if new_idx and gone:
            candidates = []
            for i in new_idx:
                for gi, (_, gshape) in enumerate(gone):
                    d = bounds_distance(shapes[i], gshape)
                    if d is not None and d < tol:
                        candidates.append((d, i, gi))
            candidates.sort()
            used_new: set[int] = set()
            used_gone: set[int] = set()
            for _, i, gi in candidates:
                if i in used_new or gi in used_gone:
                    continue
                gname = gone[gi][0]
                if gname in cur_names:
                    continue
                cur_names.discard(names[i])
                alias[shapes[i].key] = gname
                names[i] = gname
                cur_names.add(gname)
                used_new.add(i)
                used_gone.add(gi)

        out.append(names)
        prev = list(zip(names, shapes, strict=True))
    return out


def prepare_shapes(
    keyframes: Sequence[Keyframe | State], names: list[list[str]], frame_width: float
) -> list[list[Shape | None]]:
    """Decide what each slide really needs to contain.

    Two kinds of shape need care, and both are about entrances and exits:

    * **Collapsed to a point.** Many manim animations start (or end) with
      geometry squashed into a single point: ``GrowFromCenter``, ``Create``,
      ``ShrinkToCenter``... Morph can't tween out of a zero-size freeform, so
      the shape is replaced by a tiny, fully transparent copy of its
      neighbouring state, which reads as growing out of that point. Shapes
      that are being *drawn* (``Write``, ``Create`` on anything but a straight
      line) are left off instead, so Morph simply fades them.
    * **Fully transparent.** Kept only when the same shape is visible on a
      neighbouring slide somewhere else (``FadeIn(m, shift=UP)``): Morph then
      moves it while fading. Otherwise it would just be clutter.

    ``None`` marks a shape that is left off its slide.
    """
    moved = 1e-4 * frame_width
    lookup = [
        dict(zip(nm, kf.snapshot.shapes, strict=True))
        for kf, nm in zip(keyframes, names, strict=True)
    ]

    def neighbours(k: int, name: str) -> list[PathShape]:
        found = []
        for j in (k + 1, k - 1):
            if 0 <= j < len(keyframes):
                cand = lookup[j].get(name)
                if isinstance(cand, PathShape) and not cand.collapsed:
                    found.append(cand)
        return found

    out: list[list[Shape | None]] = []
    for k, (kf, nm) in enumerate(zip(keyframes, names, strict=True)):
        shapes: list[Shape | None] = []
        for shape, name in zip(kf.snapshot.shapes, nm, strict=True):
            if not isinstance(shape, PathShape):
                shapes.append(shape)
            elif shape.collapsed:
                refs = neighbours(k, name)
                if not refs:
                    shapes.append(None)  # nothing sensible to show
                elif shape.key in kf.fade_keys and not is_straight_line(
                    refs[0].subpaths, refs[0].closed
                ):
                    shapes.append(None)
                else:
                    shapes.append(_seed(shape, refs[0]))
            elif not (shape.has_fill or shape.has_stroke):
                keep = any(
                    r.visible and (bounds_distance(shape, r) or 0.0) > moved
                    for r in neighbours(k, name)
                )
                shapes.append(shape if keep else None)
            else:
                shapes.append(shape)
        out.append(shapes)
    return out


def _seed(point: PathShape, ref: PathShape) -> PathShape:
    center = np.concatenate(point.subpaths).mean(axis=0)
    return replace(
        ref.faded(),
        key=point.key,
        subpaths=[center + (sp - center) * SEED_SCALE for sp in ref.subpaths],
        stroke_width=ref.stroke_width * SEED_SCALE,  # a true miniature, outline included
    )


# --------------------------------------------------------------------------
# Placement: where each shape sits, how it is rotated, which preset it is
# --------------------------------------------------------------------------


class _Coords:
    """Maps scene units onto the slide, letterboxing if the aspects differ."""

    def __init__(self, frame_width: float, frame_height: float, slide_w: int, slide_h: int):
        self.fw, self.fh = frame_width, frame_height
        self.s = min(slide_w / frame_width, slide_h / frame_height)
        self.ox = (slide_w - frame_width * self.s) / 2
        self.oy = (slide_h - frame_height * self.s) / 2

    def x(self, x: float) -> float:
        return self.ox + (x + self.fw / 2) * self.s

    def y(self, y: float) -> float:
        return self.oy + (self.fh / 2 - y) * self.s

    def length(self, d: float) -> float:
        return d * self.s

    def points(self, pts: np.ndarray) -> np.ndarray:
        return np.column_stack(
            [
                self.ox + (pts[:, 0] + self.fw / 2) * self.s,
                self.oy + (self.fh / 2 - pts[:, 1]) * self.s,
            ]
        )


@dataclass
class Placement:
    points: list[np.ndarray]  # control points on the slide, in EMU
    closed: list[bool]
    theta: float = 0.0  # rotation on the slide
    kind: Preset | None = None  # the preset this outline happens to be
    # The same outline in the frame where it *is* a preset. Set when a shape
    # was just bent into a new outline: the rotation it carried over no longer
    # means anything, so it may start afresh once the Morph into it is done.
    rebased: Placement | None = None

    @property
    def leaving(self) -> Placement:
        return self.rebased or self

    def frame(self, parent_theta: float = 0.0) -> LocalFrame:
        """The shape's box, inside a parent that is itself rotated."""
        if parent_theta:
            inv = rotation(-parent_theta).T
            return to_local([p @ inv for p in self.points], self.theta - parent_theta)
        return to_local(self.points, self.theta)


@dataclass
class Item:
    """One shape as it is written on one slide."""

    shape: Shape
    name: str
    placement: Placement | None = None  # None for pictures
    preset: Preset | None = None
    group: str | None = None  # the PowerPoint group it is written inside
    group_theta: float = 0.0


def plan_placements(
    shapes_per_state: list[list[Shape | None]],
    names_per_state: list[list[str]],
    coords: _Coords,
    native: bool = True,
) -> list[list[Placement | None]]:
    """Give every path shape its slide coordinates, a rotation and a kind.

    Rotation is tracked state to state: when a shape's points moved rigidly,
    the turn is added to its rotation rather than baked into its geometry.
    """
    out: list[list[Placement | None]] = []
    last: dict[str, tuple[int, Placement]] = {}
    for k, (shapes, names) in enumerate(zip(shapes_per_state, names_per_state, strict=True)):
        row: list[Placement | None] = []
        for shape, name in zip(shapes, names, strict=True):
            if not isinstance(shape, PathShape):
                row.append(None)
                continue
            keep = [i for i, sp in enumerate(shape.subpaths) if len(sp) >= 4]
            if not keep:
                row.append(None)
                continue
            placement = Placement(
                points=[coords.points(shape.subpaths[i]) for i in keep],
                closed=[shape.closed[i] for i in keep],
            )
            row.append(placement)
            if not native:
                continue
            prev = last.get(name)
            single = len(placement.points) == 1
            if prev is not None and prev[0] == k - 1:
                turn = similarity_angle(prev[1].points, placement.points)
                placement.theta = prev[1].theta + (turn or 0.0)
            else:
                turn = 0.0
                placement.theta = natural_angle(placement.points, placement.closed)
            if single:
                placement.kind = classify(placement.frame(), placement.closed)
            last[name] = (k, placement)
            if single and turn is None and placement.kind is None:
                theta = natural_angle(placement.points, placement.closed)
                fresh = Placement(placement.points, placement.closed, theta)
                fresh.kind = classify(fresh.frame(), fresh.closed)
                if fresh.kind is not None:
                    placement.rebased = fresh
                    last[name] = (k, fresh)
        out.append(row)
    return out


def _first(*flags: bool | None, default: bool) -> bool:
    return next((f for f in flags if f is not None), default)


def plan_renders(
    morph: list[bool],
    shapes_per_state: list[list[Shape | None]],
    names_per_state: list[list[str]],
    placements: list[list[Placement | None]],
    group: bool = True,
) -> list[tuple[list[Item], list[Item]]]:
    """Choose, for every state, the form of each shape: once as the state is
    arrived at, and once as it is left.

    * **Preset or freeform.** Across a Morph, a shape is a preset if it is the
      same kind of preset at both ends (Morph then tweens a real rectangle,
      ellipse...). If the outline changes kind, both ends are freeforms with
      manim's matching points, so the outline bends the way manim bends it.
    * **Group or loose parts.** A compound mobject whose parts all move as one
      rigid unit across a Morph is a group at both ends. If its parts move
      relative to each other (a ``Transform``, letters arriving one by one),
      they are separate shapes at both ends so each can be tweened.

    A side with no Morph (the first slide, a hard cut) simply takes the form
    of the other side.
    """
    n = len(shapes_per_state)

    def entries(i: int) -> dict[str, tuple[int, PathShape, Placement]]:
        return {
            name: (j, shape, pl)
            for j, (shape, name, pl) in enumerate(
                zip(shapes_per_state[i], names_per_state[i], placements[i], strict=True)
            )
            if isinstance(shape, PathShape) and pl is not None
        }

    index = [entries(i) for i in range(n)]

    # Compound mobjects on each state: group name -> member names, in paint
    # order. A group whose members are not painted consecutively cannot be a
    # PowerPoint group without changing what overlaps what.
    members: list[dict[str, list[str]]] = []
    for i in range(n):
        found: dict[str, list[str]] = {}
        broken: set[str] = set()
        previous: str | None = None
        for shape, name, pl in zip(
            shapes_per_state[i], names_per_state[i], placements[i], strict=True
        ):
            if shape is None:
                continue
            g = shape.group if isinstance(shape, PathShape) and pl is not None else None
            if g is not None:
                if g in found and previous != g:
                    broken.add(g)
                found.setdefault(g, []).append(name)
            previous = g
        members.append({g: m for g, m in found.items() if g not in broken} if group else {})

    def points_of(i: int, names: list[str]) -> list[np.ndarray]:
        return [p for name in names for p in index[i][name][2].points]

    # What each Morph asks of the shapes it connects.
    same_kind: list[dict[str, bool]] = []
    as_unit: list[dict[str, bool]] = []
    group_theta: list[dict[str, float]] = [{} for _ in range(n)]
    for g, names in members[0].items() if n else ():
        group_theta[0][g] = index[0][names[0]][2].theta if len(names) == 1 else 0.0
    for i in range(n - 1):
        kinds: dict[str, bool] = {}
        units: dict[str, bool] = {}
        for name, (_, _, a) in index[i].items():
            if name in index[i + 1]:
                a, b = a.leaving, index[i + 1][name][2]
                kinds[name] = (
                    a.kind is not None and b.kind is not None and a.kind.kind == b.kind.kind
                )
        for g, names in members[i + 1].items():
            before = members[i].get(g)
            move = None
            if before == names:
                move = similarity(points_of(i, names), points_of(i + 1, names))
            if before == names:
                turn = float(np.angle(move)) if move is not None else 0.0
                group_theta[i + 1][g] = group_theta[i][g] + turn
            else:
                group_theta[i + 1][g] = index[i + 1][names[0]][2].theta if len(names) == 1 else 0.0
            if before is None:
                continue
            if len(names) > 1:
                units[g] = move is not None
            else:
                # A lone shape needs the wrapper only where it is a freeform
                # that changes size; everything else Morph handles cleanly.
                resized = move is not None and abs(abs(move) - 1.0) > 2e-3
                units[g] = resized and not kinds.get(names[0], False)
        same_kind.append(kinds if morph[i] else {})
        as_unit.append(units if morph[i] else {})

    renders: list[tuple[list[Item], list[Item]]] = []
    for i in range(n):
        kind_in = same_kind[i - 1] if i > 0 else {}
        kind_out = same_kind[i] if i < n - 1 else {}
        unit_in = as_unit[i - 1] if i > 0 else {}
        unit_out = as_unit[i] if i < n - 1 else {}
        sides: list[list[Item]] = []
        for leaving, first_kind, then_kind, first_unit, then_unit in (
            (False, kind_in, kind_out, unit_in, unit_out),
            (True, kind_out, kind_in, unit_out, unit_in),
        ):
            grouped = {
                g
                for g, names in members[i].items()
                if _first(first_unit.get(g), then_unit.get(g), default=len(names) > 1)
            }
            items: list[Item] = []
            for shape, name, pl in zip(
                shapes_per_state[i], names_per_state[i], placements[i], strict=True
            ):
                if shape is None:
                    continue
                if not isinstance(shape, PathShape):
                    items.append(Item(shape, name))
                    continue
                if pl is None:
                    continue
                if leaving:
                    pl = pl.leaving
                item = Item(shape, name, pl)
                if _first(first_kind.get(name), then_kind.get(name), default=pl.kind is not None):
                    item.preset = pl.kind  # None if this outline is no preset after all
                if shape.group in grouped:
                    item.group = shape.group
                    item.group_theta = group_theta[i][shape.group]
                items.append(item)
            sides.append(items)
        renders.append((sides[0], sides[1]))
    return renders


def merge_sides(arrive: list[Item], depart: list[Item], tol: float) -> list[list[Item]]:
    """One slide if the two sides of a keyframe can share it, else two.

    They can share a slide when every shape they have in common takes the
    same form and the same outline on both. Shapes that only the arriving
    side has (the last, invisible state of something that just left) are
    carried over.
    """
    leaving = {item.name: item for item in depart}
    for a in arrive:
        b = leaving.get(a.name)
        if b is None:
            continue
        same_form = (
            (a.preset.kind if a.preset else None) == (b.preset.kind if b.preset else None)
            and a.group == b.group
            and abs(a.group_theta - b.group_theta) < 1e-6
        )
        if not same_form:
            return [arrive, depart]
        if a.placement is None or b.placement is None or a.placement is b.placement:
            continue
        if not (a.shape.visible and b.shape.visible):
            continue
        if abs(a.placement.theta - b.placement.theta) > 1e-6:
            return [arrive, depart]
        pa, pb = a.placement.points, b.placement.points
        same_outline = len(pa) == len(pb) and all(
            x.shape == y.shape and np.allclose(x, y, atol=tol) for x, y in zip(pa, pb, strict=True)
        )
        if not same_outline and a.preset is None:
            return [arrive, depart]

    extras: dict[str | None, list[Item]] = {}
    anchor: str | None = None
    for a in arrive:
        if a.name in leaving:
            anchor = a.name
        else:
            extras.setdefault(anchor, []).append(a)
    if not extras:
        return [depart]
    merged = list(extras.get(None, []))
    for b in depart:
        merged.append(b)
        merged.extend(extras.get(b.name, []))
    return [merged]


# --------------------------------------------------------------------------
# DrawingML generation
# --------------------------------------------------------------------------


def _hex(rgb: Any) -> str:
    r, g, b = (max(0, min(255, round(c * 255))) for c in rgb[:3])
    return f"{r:02X}{g:02X}{b:02X}"


def _color(rgba: Any) -> str:
    alpha = ""
    if len(rgba) > 3 and rgba[3] < 0.999:
        alpha = f'<a:alpha val="{max(0, round(rgba[3] * 100000))}"/>'
    return f'<a:srgbClr val="{_hex(rgba)}">{alpha}</a:srgbClr>'


def _paint(color: Any, gradient: Gradient | None, theta: float) -> str:
    """A solid or gradient fill element."""
    if gradient is None:
        return f"<a:solidFill>{_color(color)}</a:solidFill>"
    n = len(gradient.stops)
    stops = "".join(
        f'<a:gs pos="{round(i * 100000 / (n - 1))}">{_color(c)}</a:gs>'
        for i, c in enumerate(gradient.stops)
    )
    # Scene y is up, slide y is down; the angle is relative to the shape's
    # own (possibly rotated) frame because rotWithShape is on.
    dx, dy = gradient.direction
    ang = (math.degrees(math.atan2(-dy, dx) - theta)) % 360
    return (
        f'<a:gradFill rotWithShape="1"><a:gsLst>{stops}</a:gsLst>'
        f'<a:lin ang="{round(ang * 60000) % 21600000}" scaled="0"/></a:gradFill>'
    )


def _rot_attr(theta: float) -> str:
    rot = round(math.degrees(theta) * 60000) % 21600000
    return f' rot="{rot}"' if rot else ""


def _custom_geometry(
    frame: LocalFrame, closed: list[bool], w: int, h: int, fill: bool, stroke: bool
) -> str:
    segs = []
    for pts, is_closed in zip(frame.subpaths, closed, strict=True):
        p = np.rint(pts).astype(int)
        segs.append(f'<a:moveTo><a:pt x="{p[0, 0]}" y="{p[0, 1]}"/></a:moveTo>')
        for i in range(0, len(p) - 3, 4):
            segs.append(
                "<a:cubicBezTo>"
                f'<a:pt x="{p[i + 1, 0]}" y="{p[i + 1, 1]}"/>'
                f'<a:pt x="{p[i + 2, 0]}" y="{p[i + 2, 1]}"/>'
                f'<a:pt x="{p[i + 3, 0]}" y="{p[i + 3, 1]}"/>'
                "</a:cubicBezTo>"
            )
        if is_closed:
            segs.append("<a:close/>")
    attrs = f'w="{w}" h="{h}"'
    if not fill:
        attrs += ' fill="none"'
    if not stroke:
        attrs += ' stroke="0"'
    return (
        "<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>"
        '<a:rect l="0" t="0" r="r" b="b"/>'
        f"<a:pathLst><a:path {attrs}>{''.join(segs)}</a:path></a:pathLst>"
        "</a:custGeom>"
    )


def _preset_geometry(preset: Preset) -> str:
    adj = ""
    if preset.adj is not None:
        adj = f'<a:gd name="adj" fmla="val {round(preset.adj * 100000)}"/>'
    return f'<a:prstGeom prst="{preset.kind}"><a:avLst>{adj}</a:avLst></a:prstGeom>'


def path_shape_xml(item: Item, name: str, shape_id: int, coords: _Coords) -> str:
    """Build the ``<p:sp>`` (or ``<p:cxnSp>`` for a line) for one VMobject."""
    shape, placement, preset = item.shape, item.placement, item.preset
    assert isinstance(shape, PathShape) and placement is not None
    frame = placement.frame(item.group_theta if item.group else 0.0)
    line_w = round(coords.length(shape.stroke_width * MANIM_STROKE_UNIT))
    # An outline-only shape gets a true "No Fill" (and a filled, unstroked one
    # "No Line"). Only a shape that is invisible altogether keeps both paints
    # at zero opacity, for Morph to fade up from.
    ghost = not (shape.has_fill or shape.has_stroke)
    has_fill = (shape.has_fill or ghost) and (
        shape.fill is not None or shape.fill_gradient is not None
    )
    has_stroke = (
        (shape.has_stroke or ghost)
        and (shape.stroke is not None or shape.stroke_gradient is not None)
        and line_w >= 1
    )
    is_line = preset is not None and preset.kind == "line"

    w = max(1, round(frame.w))
    h = 0 if is_line else max(1, round(frame.h))
    xfrm = (
        f"<a:xfrm{_rot_attr(frame.theta)}>"
        f'<a:off x="{round(frame.x)}" y="{round(frame.y)}"/><a:ext cx="{w}" cy="{h}"/>'
        "</a:xfrm>"
    )
    if preset is not None:
        geom = _preset_geometry(preset)
    else:
        geom = _custom_geometry(frame, placement.closed, w, h, has_fill, has_stroke)

    # Gradients turn with the shape, so their angle is measured against the
    # shape's total rotation on the slide.
    if has_fill and not is_line:
        fill_xml = _paint(shape.fill, shape.fill_gradient, placement.theta)
    else:
        fill_xml = "<a:noFill/>"
    if has_stroke:
        join = {"round": "<a:round/>", "bevel": "<a:bevel/>"}.get(
            shape.join, '<a:miter lim="800000"/>'
        )
        ln_xml = (
            f'<a:ln w="{line_w}" cap="{shape.cap}">'
            f"{_paint(shape.stroke, shape.stroke_gradient, placement.theta)}{join}</a:ln>"
        )
    else:
        ln_xml = "<a:ln><a:noFill/></a:ln>"

    c_nv_pr = f'<p:cNvPr id="{shape_id}" name="{_quote(name)}"/>'
    if is_line:
        return (
            f"<p:cxnSp {_NS}>"
            f"<p:nvCxnSpPr>{c_nv_pr}<p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>"
            f"<p:spPr>{xfrm}{geom}{ln_xml}</p:spPr>"
            "</p:cxnSp>"
        )
    return (
        f"<p:sp {_NS}>"
        f"<p:nvSpPr>{c_nv_pr}<p:cNvSpPr/><p:nvPr/></p:nvSpPr>"
        f"<p:spPr>{xfrm}{geom}{fill_xml}{ln_xml}</p:spPr>"
        '<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/>'
        '<a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US"/></a:p></p:txBody>'
        "</p:sp>"
    )


def group_xml(items: list[Item], first_id: int, coords: _Coords) -> str:
    """A ``<p:grpSp>`` holding the parts of one compound mobject.

    Morph pairs the group (not its children) across slides, so only the group
    carries the ``!!`` name.
    """
    name, theta = items[0].group, items[0].group_theta
    inv = rotation(-theta).T
    local = np.concatenate([p @ inv for item in items for p in item.placement.points])  # type: ignore[union-attr]
    lo, hi = local.min(axis=0), local.max(axis=0)
    w, h = (max(1, round(v)) for v in hi - lo)
    center = rotation(theta) @ ((lo + hi) / 2)
    children = "".join(
        path_shape_xml(item, item.name, first_id + 1 + n, coords) for n, item in enumerate(items)
    )
    return (
        f"<p:grpSp {_NS}>"
        f'<p:nvGrpSpPr><p:cNvPr id="{first_id}" name="{_quote(MORPH_PREFIX + str(name))}"/>'
        "<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"
        f"<p:grpSpPr><a:xfrm{_rot_attr(theta)}>"
        f'<a:off x="{round(center[0] - w / 2)}" y="{round(center[1] - h / 2)}"/>'
        f'<a:ext cx="{w}" cy="{h}"/>'
        f'<a:chOff x="{round(lo[0])}" y="{round(lo[1])}"/><a:chExt cx="{w}" cy="{h}"/>'
        "</a:xfrm></p:grpSpPr>"
        f"{children}</p:grpSp>"
    )


def _quote(text: str) -> str:
    return escape(text, {'"': "&quot;"})


def transition_xml(duration_ms: int, advance_after_ms: int | None) -> etree._Element:
    """Morph transition into a slide, with a fade fallback for old viewers."""
    adv = f' advTm="{advance_after_ms}"' if advance_after_ms is not None else ""
    if duration_ms <= 0:
        # Hard cut: a transition element is only needed to carry advTm.
        xml = f"<p:transition {nsdecls('p')}{adv}/>"
        return parse_xml(xml)
    xml = (
        f'<mc:AlternateContent xmlns:mc="{NS_MC}">'
        f'<mc:Choice xmlns:p159="{NS_P159}" xmlns:p14="{NS_P14}" Requires="p159">'
        f'<p:transition {nsdecls("p")} spd="slow" p14:dur="{duration_ms}"{adv}>'
        f'<p159:morph option="byObject"/></p:transition>'
        f"</mc:Choice>"
        f"<mc:Fallback>"
        f'<p:transition {nsdecls("p")} spd="fast"{adv}><p:fade/></p:transition>'
        f"</mc:Fallback>"
        f"</mc:AlternateContent>"
    )
    return parse_xml(xml)


# --------------------------------------------------------------------------
# Presentation assembly
# --------------------------------------------------------------------------


def advance_after(timeline: Timeline, i: int) -> int | None:
    """Milliseconds after which slide ``i`` auto-advances (None = on click)."""
    kfs = timeline.keyframes
    if i == len(kfs) - 1:
        return None
    nxt = kfs[i + 1]
    if timeline.mode == "click":
        return None if nxt.starts_play else 0
    if kfs[i].marker:
        return None
    return max(0, kfs[i].hold_ms)


def write_pptx(
    timeline: Timeline,
    path: str | Path,
    *,
    template: str | Path | None = None,
    background: bool | None = None,
    native_shapes: bool = True,
    group: bool = True,
) -> Path:
    """Write ``timeline`` to ``path`` and return the path.

    Use :func:`write_deck` to also learn how many slides were written.
    """
    return write_deck(
        timeline,
        path,
        template=template,
        background=background,
        native_shapes=native_shapes,
        group=group,
    )[0]


def write_deck(
    timeline: Timeline,
    path: str | Path,
    *,
    template: str | Path | None = None,
    background: bool | None = None,
    native_shapes: bool = True,
    group: bool = True,
) -> tuple[Path, int]:
    """Write ``timeline`` to ``path``; return the path and the slide count.

    Parameters
    ----------
    template
        A ``.pptx`` to build on. Its theme, slide size and layouts are used,
        any slides it already contains are kept, and the animation is appended
        after them. Without a template a clean 16:9 deck is created.
    background
        Paint the scene's background color onto the slides. Defaults to
        ``True`` for a fresh deck and ``False`` with a template, where the
        template's own background shows through instead.
    native_shapes
        Write rectangles, ellipses, lines... as PowerPoint presets and store
        rotation as a real rotation. Turn off to get plain freeforms only.
    group
        Write compound mobjects (a word, an equation, a set of axes) that only
        ever move as a whole as one PowerPoint group. See :func:`plan_renders`.
    """
    path = Path(path)
    kfs = timeline.keyframes
    paint_background = background if background is not None else template is None

    prs = Presentation(str(template)) if template is not None else _new_presentation(timeline)
    existing = len(prs.slides)
    coords = _Coords(timeline.frame_width, timeline.frame_height, prs.slide_width, prs.slide_height)
    layout = _blank_layout(prs)

    states, morph = expand(kfs)
    names = resolve_names(states, timeline.frame_width)
    shapes = prepare_shapes(states, names, timeline.frame_width)
    placements = plan_placements(shapes, names, coords, native=native_shapes)
    renders = plan_renders(morph, shapes, names, placements, group=group)
    first_state = {state.keyframe: i for i, state in reversed(list(enumerate(states)))}
    last_state = {state.keyframe: i for i, state in enumerate(states)}

    backgrounds = [_hex(kf.snapshot.background) for kf in kfs]
    master_background = None
    if paint_background and template is None and backgrounds:
        master_background = Counter(backgrounds).most_common(1)[0][0]
        _set_master_background(prs, master_background)

    section_starts: list[tuple[str | None, int]] = []
    for i, kf in enumerate(kfs):
        variants = merge_sides(
            renders[first_state[i]][0], renders[last_state[i]][1], tol=1e-5 * prs.slide_width
        )
        for v, items in enumerate(variants):
            slide = prs.slides.add_slide(layout)
            if v == 0:
                section_starts.append((kf.section, len(prs.slides) - 1))
            for placeholder in list(slide.placeholders):
                placeholder._element.getparent().remove(placeholder._element)
            if paint_background and backgrounds[i] != master_background:
                fill = slide.background.fill
                fill.solid()
                fill.fore_color.rgb = RGBColor.from_string(backgrounds[i])
            _add_shapes(slide, items, coords)

            # A keyframe written twice: Morph into the first copy, which is
            # left the way the keyframe is (click or timer); the second copy
            # replaces it instantly and moves straight on.
            duration = kf.transition_ms if (i > 0 and v == 0) else 0
            advance = advance_after(timeline, i) if v == 0 else 0
            trans = transition_xml(duration, advance)
            clr = slide._element.find(qn("p:clrMapOvr"))
            if clr is not None:
                clr.addnext(trans)
            else:
                slide._element.append(trans)

            if kf.notes:
                slide.notes_slide.notes_text_frame.text = kf.notes

    _add_sections(prs, timeline.title, section_starts, existing)
    if template is None:
        _set_properties(prs, timeline)

    path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(path))
    return path, len(prs.slides) - existing


def _new_presentation(timeline: Timeline) -> Any:
    """A blank 16:9 deck whose master and layouts actually fit 16:9."""
    prs = Presentation()
    old_w, old_h = prs.slide_width, prs.slide_height
    slide_w = SLIDE_WIDTH_EMU
    slide_h = round(slide_w * timeline.frame_height / timeline.frame_width)
    prs.slide_width, prs.slide_height = Emu(slide_w), Emu(slide_h)
    # python-pptx's built-in template is 4:3 and says so.
    sld_sz = prs.element.find(qn("p:sldSz"))
    if sld_sz is not None and "type" in sld_sz.attrib:
        del sld_sz.attrib["type"]
    # Stretch the placeholders too, so that slides the user adds by hand
    # afterwards ("Title and Content"...) are laid out for the new size.
    sx, sy = slide_w / old_w, slide_h / old_h
    for part in [*prs.slide_masters, *prs.slide_layouts]:
        for xfrm in part._element.iter(qn("a:xfrm")):
            off, ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
            if off is not None:
                off.set("x", str(round(int(off.get("x")) * sx)))
                off.set("y", str(round(int(off.get("y")) * sy)))
            if ext is not None:
                ext.set("cx", str(round(int(ext.get("cx")) * sx)))
                ext.set("cy", str(round(int(ext.get("cy")) * sy)))
    return prs


def _blank_layout(prs: Any) -> Any:
    for layout in prs.slide_layouts:
        if layout.name.strip().lower() == "blank":
            return layout
    return min(prs.slide_layouts, key=lambda layout: len(layout.placeholders))


def _set_master_background(prs: Any, hex_color: str) -> None:
    """Put the scene background on the slide master.

    One place to restyle the whole deck, and slides added later match. On a
    dark background the theme's color map is flipped as well, so that text
    typed into those slides is light instead of black-on-black.
    """
    master = prs.slide_masters[0]._element
    c_sld = master.find(qn("p:cSld"))
    old = c_sld.find(qn("p:bg"))
    if old is not None:
        c_sld.remove(old)
    c_sld.insert(
        0,
        parse_xml(
            f'<p:bg {_NS}><p:bgPr><a:solidFill><a:srgbClr val="{hex_color}"/></a:solidFill>'
            "<a:effectLst/></p:bgPr></p:bg>"
        ),
    )
    r, g, b = (int(hex_color[i : i + 2], 16) / 255 for i in (0, 2, 4))
    dark = 0.2126 * r + 0.7152 * g + 0.0722 * b < 0.5
    clr_map = master.find(qn("p:clrMap"))
    if clr_map is not None:
        mapping = ("dk1", "lt1", "dk2", "lt2") if dark else ("lt1", "dk1", "lt2", "dk2")
        for attr, value in zip(("bg1", "tx1", "bg2", "tx2"), mapping, strict=True):
            clr_map.set(attr, value)


def _add_shapes(slide: Any, items: list[Item], coords: _Coords) -> None:
    tree = slide.shapes._spTree
    next_id = max((int(i) for i in tree.xpath(".//p:cNvPr/@id")), default=1) + 1
    i = 0
    while i < len(items):
        item = items[i]
        if item.group is not None:
            j = i
            while j < len(items) and items[j].group == item.group:
                j += 1
            tree.append(parse_xml(group_xml(items[i:j], next_id, coords)))
            next_id += 1 + (j - i)
            i = j
            continue
        i += 1
        shape = item.shape
        if isinstance(shape, PathShape):
            tree.append(parse_xml(path_shape_xml(item, MORPH_PREFIX + item.name, next_id, coords)))
            next_id += 1
        elif isinstance(shape, ImageShape):
            w, h = coords.length(shape.width), coords.length(shape.height)
            pic = slide.shapes.add_picture(
                io.BytesIO(shape.png),
                Emu(round(coords.x(shape.center[0]) - w / 2)),
                Emu(round(coords.y(shape.center[1]) - h / 2)),
                Emu(max(1, round(w))),
                Emu(max(1, round(h))),
            )
            pic.name = MORPH_PREFIX + item.name
            pic.rotation = -math.degrees(shape.angle) % 360
            next_id = max(next_id, pic.shape_id + 1)


def _add_sections(
    prs: Any, title: str | None, starts: list[tuple[str | None, int]], existing: int
) -> None:
    """Group slides into PowerPoint sections at each named ``next_slide()``.

    ``starts`` lists, for every keyframe, its section name (if one begins
    there) and the index of its first slide.
    """
    if not any(name for name, _ in starts):
        return
    root = prs.element
    if next(root.iter(f"{{{NS_P14}}}sectionLst"), None) is not None:
        return  # the template manages its own sections
    ids = [s.get("id") for s in root.find(qn("p:sldIdLst"))]
    sections: list[tuple[str, int]] = []  # name, first slide index
    if existing:
        sections.append(("Default Section", 0))
    for n, (name, index) in enumerate(starts):
        if name or n == 0:
            sections.append((name or title or "Animation", index))

    parts = []
    for n, (name, start) in enumerate(sections):
        end = sections[n + 1][1] if n + 1 < len(sections) else len(ids)
        guid = str(uuid.uuid5(uuid.NAMESPACE_URL, f"manim-ppt/{n}/{name}")).upper()
        slide_ids = "".join(f'<p14:sldId id="{i}"/>' for i in ids[start:end])
        parts.append(
            f'<p14:section name="{_quote(name)}" id="{{{guid}}}">'
            f"<p14:sldIdLst>{slide_ids}</p14:sldIdLst></p14:section>"
        )
    ext = parse_xml(
        f'<p:ext {nsdecls("p")} uri="{SECTION_EXT_URI}">'
        f'<p14:sectionLst xmlns:p14="{NS_P14}">{"".join(parts)}</p14:sectionLst></p:ext>'
    )
    ext_lst = root.find(qn("p:extLst"))
    if ext_lst is None:
        ext_lst = parse_xml(f"<p:extLst {nsdecls('p')}/>")
        root.append(ext_lst)
    ext_lst.append(ext)


def _set_properties(prs: Any, timeline: Timeline) -> None:
    props = prs.core_properties
    now = dt.datetime.now(dt.timezone.utc).replace(microsecond=0, tzinfo=None)
    props.title = timeline.title or ""
    props.author = ""
    props.last_modified_by = ""
    props.comments = "Exported from Manim with manim-ppt"
    props.created = now
    props.modified = now
    props.revision = 1
