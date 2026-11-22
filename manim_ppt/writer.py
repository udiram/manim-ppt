"""Turn a :class:`~manim_ppt.model.Timeline` into a .pptx built from native
shapes, chained together with PowerPoint's Morph transition.

How it works
------------
* Every keyframe becomes one slide.
* Every manim VMobject becomes a freeform shape (``a:custGeom``) whose path
  is made of ``a:cubicBezTo`` segments, which are exactly manim's cubic Beziers.
* A shape keeps the same name on every slide (``!!<key>``). The ``!!`` prefix
  tells Morph to pair shapes by name, so it tweens position, geometry, color
  and opacity between slides instead of cross-fading.
* The slide-advance flags reproduce the scene's timing: sub-steps of one
  ``play()`` advance automatically, and a click is needed between plays (or
  only at ``next_slide()`` markers).
"""

from __future__ import annotations

import io
from pathlib import Path
from xml.sax.saxutils import escape

import numpy as np
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Emu

from .model import ImageShape, Keyframe, PathShape, Snapshot, Timeline

SLIDE_WIDTH_EMU = 12_192_000  # 13.333 in, the PowerPoint 16:9 default
MANIM_STROKE_UNIT = 0.01  # cairo: line width = stroke_width * 0.01 scene units

NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
NS_P159 = "http://schemas.microsoft.com/office/powerpoint/2015/09/main"


# --------------------------------------------------------------------------
# Shape naming across slides
# --------------------------------------------------------------------------


def resolve_names(keyframes: list[Keyframe], frame_width: float) -> list[list[str]]:
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
    prev: list[tuple[str, object]] = []

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
                for gi, (gname, gshape) in enumerate(gone):
                    d = _shape_distance(shapes[i], gshape)
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
        prev = list(zip(names, shapes))
    return out


def seed_degenerate(
    keyframes: list[Keyframe], names: list[list[str]], frame_width: float
) -> list[list[object]]:
    """Replace shapes that have collapsed to a point with a tiny copy of
    their neighbouring state.

    Many manim animations start (or end) with geometry squashed into a single
    point: ``GrowFromCenter``, ``Create``, ``Write``, ``ShrinkToCenter``...
    Morph can't tween out of a zero-size freeform (it draws stray lines), but
    it tweens nicely from a 2%-scale copy of the same path placed at that
    point, which reads as the object growing out of it.
    """
    eps = 1e-3 * frame_width
    out: list[list[object]] = []
    lookup = [
        {n: s for n, s in zip(nm, kf.snapshot.shapes)} for kf, nm in zip(keyframes, names)
    ]
    for k, (kf, nm) in enumerate(zip(keyframes, names)):
        shapes: list[object] = []
        for shape, name in zip(kf.snapshot.shapes, nm):
            if isinstance(shape, PathShape) and _is_degenerate(shape, eps):
                ref = None
                for j in (k + 1, k - 1):
                    if 0 <= j < len(keyframes):
                        cand = lookup[j].get(name)
                        if isinstance(cand, PathShape) and not _is_degenerate(cand, eps):
                            ref = cand
                            break
                if ref is None:
                    shapes.append(None)  # nothing sensible to show
                    continue
                center = np.concatenate(shape.subpaths).mean(axis=0)
                shape = PathShape(
                    key=shape.key,
                    subpaths=[center + (sp - center) * 0.02 for sp in ref.subpaths],
                    closed=list(ref.closed),
                    fill=shape.fill,
                    stroke=shape.stroke,
                    stroke_width=shape.stroke_width,
                    cap=shape.cap,
                    join=shape.join,
                )
            shapes.append(shape)
        out.append(shapes)
    return out


def _is_degenerate(shape: PathShape, eps: float) -> bool:
    x0, y0, x1, y1 = shape.bounds()
    return max(x1 - x0, y1 - y0) < eps


def _shape_distance(a, b) -> float | None:
    if type(a) is not type(b):
        return None
    ax0, ay0, ax1, ay1 = a.bounds()
    bx0, by0, bx1, by1 = b.bounds()
    return max(abs(ax0 - bx0), abs(ay0 - by0), abs(ax1 - bx1), abs(ay1 - by1))


# --------------------------------------------------------------------------
# DrawingML generation
# --------------------------------------------------------------------------


class _Coords:
    def __init__(self, frame_width: float, frame_height: float, slide_w: int, slide_h: int):
        self.fw, self.fh = frame_width, frame_height
        self.s = slide_w / frame_width
        self.slide_w, self.slide_h = slide_w, slide_h

    def x(self, x: float) -> float:
        return (x + self.fw / 2) * self.s

    def y(self, y: float) -> float:
        return (self.fh / 2 - y) * self.s

    def length(self, d: float) -> float:
        return d * self.s


def _hex(rgb) -> str:
    r, g, b = (max(0, min(255, round(c * 255))) for c in rgb[:3])
    return f"{r:02X}{g:02X}{b:02X}"


def _color(rgba) -> str:
    alpha = ""
    if len(rgba) > 3 and rgba[3] < 0.999:
        alpha = f'<a:alpha val="{max(0, round(rgba[3] * 100000))}"/>'
    return f'<a:srgbClr val="{_hex(rgba)}">{alpha}</a:srgbClr>'


def path_shape_xml(shape: PathShape, name: str, shape_id: int, c: _Coords) -> str | None:
    """Build a ``<p:sp>`` freeform for one manim VMobject."""
    fill = shape.fill if shape.fill and shape.fill[3] > 0.001 else None
    line_w = c.length(shape.stroke_width * MANIM_STROKE_UNIT)
    stroke = shape.stroke if (shape.stroke and shape.stroke[3] > 0.001 and line_w >= 1) else None
    if fill is None and stroke is None:
        return None

    # Transform every control point into slide EMUs.
    subpaths = []
    for sp in shape.subpaths:
        if len(sp) < 4:
            continue
        pts = np.column_stack([(sp[:, 0] + c.fw / 2) * c.s, (c.fh / 2 - sp[:, 1]) * c.s])
        subpaths.append(pts)
    if not subpaths:
        return None
    allpts = np.concatenate(subpaths)
    x0, y0 = allpts.min(axis=0)
    x1, y1 = allpts.max(axis=0)
    off_x, off_y = int(round(x0)), int(round(y0))
    w = max(1, int(round(x1 - off_x)))
    h = max(1, int(round(y1 - off_y)))

    def pt(p) -> str:
        return f'<a:pt x="{int(round(p[0] - off_x))}" y="{int(round(p[1] - off_y))}"/>'

    segs = []
    for pts, closed in zip(subpaths, shape.closed):
        segs.append(f"<a:moveTo>{pt(pts[0])}</a:moveTo>")
        for i in range(0, len(pts) - 3, 4):
            segs.append(f"<a:cubicBezTo>{pt(pts[i + 1])}{pt(pts[i + 2])}{pt(pts[i + 3])}</a:cubicBezTo>")
        if closed:
            segs.append("<a:close/>")

    path_attrs = f'w="{w}" h="{h}"'
    if fill is None:
        path_attrs += ' fill="none"'
    if stroke is None:
        path_attrs += ' stroke="0"'

    fill_xml = f"<a:solidFill>{_color(fill)}</a:solidFill>" if fill else "<a:noFill/>"
    if stroke:
        join = {"round": "<a:round/>", "bevel": "<a:bevel/>"}.get(shape.join, '<a:miter lim="800000"/>')
        ln_xml = (
            f'<a:ln w="{int(round(line_w))}" cap="{shape.cap}">'
            f"<a:solidFill>{_color(stroke)}</a:solidFill>{join}</a:ln>"
        )
    else:
        ln_xml = "<a:ln><a:noFill/></a:ln>"

    return (
        f"<p:sp {nsdecls('a', 'p')}>"
        f'<p:nvSpPr><p:cNvPr id="{shape_id}" name="{escape(name, {chr(34): "&quot;"})}"/>'
        f"<p:cNvSpPr/><p:nvPr/></p:nvSpPr>"
        f"<p:spPr>"
        f'<a:xfrm><a:off x="{off_x}" y="{off_y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
        f"<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>"
        f'<a:rect l="0" t="0" r="r" b="b"/>'
        f"<a:pathLst><a:path {path_attrs}>{''.join(segs)}</a:path></a:pathLst>"
        f"</a:custGeom>"
        f"{fill_xml}{ln_xml}"
        f"</p:spPr>"
        f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>'
        f"</p:sp>"
    )


def transition_xml(duration_ms: int, advance_after_ms: int | None) -> etree._Element:
    """Morph transition into a slide, with a fade fallback for old viewers."""
    adv = f' advTm="{advance_after_ms}"' if advance_after_ms is not None else ""
    if duration_ms <= 0:
        # Hard cut: a transition element is only needed to carry advTm.
        xml = f'<p:transition {nsdecls("p")}{adv}/>'
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


def write_pptx(timeline: Timeline, path: str | Path) -> Path:
    path = Path(path)
    prs = Presentation()
    slide_w = SLIDE_WIDTH_EMU
    slide_h = int(round(slide_w * timeline.frame_height / timeline.frame_width))
    prs.slide_width, prs.slide_height = Emu(slide_w), Emu(slide_h)
    coords = _Coords(timeline.frame_width, timeline.frame_height, slide_w, slide_h)
    blank = prs.slide_layouts[6]

    names_per_slide = resolve_names(timeline.keyframes, timeline.frame_width)
    shapes_per_slide = seed_degenerate(timeline.keyframes, names_per_slide, timeline.frame_width)

    for i, (kf, names, shapes) in enumerate(
        zip(timeline.keyframes, names_per_slide, shapes_per_slide)
    ):
        slide = prs.slides.add_slide(blank)
        bg = slide.background.fill
        bg.solid()
        bg.fore_color.rgb = RGBColor.from_string(_hex(kf.snapshot.background))
        _add_shapes(slide, shapes, names, coords)

        trans = transition_xml(kf.transition_ms if i > 0 else 0, advance_after(timeline, i))
        clr = slide._element.find(qn("p:clrMapOvr"))
        if clr is not None:
            clr.addnext(trans)
        else:
            slide._element.append(trans)

        if kf.notes:
            slide.notes_slide.notes_text_frame.text = kf.notes

    path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(path)
    return path


def _add_shapes(slide, shapes: list, names: list[str], coords: _Coords) -> None:
    tree = slide.shapes._spTree
    next_id = 2
    for shape, name in zip(shapes, names):
        if shape is None:
            continue
        morph_name = f"!!{name}"
        if isinstance(shape, PathShape):
            xml = path_shape_xml(shape, morph_name, next_id, coords)
            if xml is None:
                continue
            tree.append(parse_xml(xml))
            next_id += 1
        elif isinstance(shape, ImageShape):
            x0, y0, x1, y1 = shape.box
            left, top = coords.x(x0), coords.y(y1)
            pic = slide.shapes.add_picture(
                io.BytesIO(shape.png),
                Emu(int(round(left))),
                Emu(int(round(top))),
                Emu(max(1, int(round(coords.length(x1 - x0))))),
                Emu(max(1, int(round(coords.length(y1 - y0))))),
            )
            pic.name = morph_name
            next_id = max(next_id, pic.shape_id + 1)

