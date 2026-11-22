"""Tests for the manim-independent half: naming, shape forms, slide flow."""

import io
import re

import numpy as np
import pytest
from PIL import Image
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Emu

from manim_ppt.model import Gradient, ImageShape, Keyframe, PathShape, Snapshot, Timeline
from manim_ppt.writer import advance_after, prepare_shapes, resolve_names, write_pptx

from .shapes import FW, SLIDE_H, SLIDE_W, blob, circle, polygon, square

UNIT = SLIDE_W / FW  # EMU per scene unit


def deck(tmp_path, keyframes, mode="click", **options):
    """Write keyframes and return the slides of the resulting deck."""
    for kf in keyframes[1:]:
        if kf.transition_ms == 0:
            kf.transition_ms = 1000
            kf.starts_play = True
    out = write_pptx(Timeline(keyframes=keyframes, mode=mode), tmp_path / "t.pptx", **options)
    return list(Presentation(out).slides)


def kf(*shapes, **kwargs):
    return Keyframe(Snapshot(list(shapes)), **kwargs)


def tree(slide):
    return slide.shapes._spTree


def top_names(slide):
    return [
        el.xpath("./*/p:cNvPr/@name")[0]
        for el in tree(slide)
        if el.tag.endswith(("}sp", "}cxnSp", "}grpSp", "}pic"))
    ]


def find(slide, name):
    hits = tree(slide).xpath(f'.//*[*/p:cNvPr/@name="{name}"]')
    assert hits, f"{name} not on slide: {top_names(slide)}"
    return hits[0]


def xfrm(el):
    x = el.find(".//" + qn("a:xfrm"))
    off, ext = x.find(qn("a:off")), x.find(qn("a:ext"))
    return {
        "rot": int(x.get("rot", 0)) / 60000,
        "x": int(off.get("x")),
        "y": int(off.get("y")),
        "w": int(ext.get("cx")),
        "h": int(ext.get("cy")),
    }


def preset(el):
    found = el.xpath("./p:spPr/a:prstGeom/@prst")
    return found[0] if found else None


def is_morph(slide):
    return "p159:morph" in slide._element.xml


def transition(slide):
    return re.findall(r"<p:transition[^>]*>", slide._element.xml)[0]


# ---------------------------------------------------------------- shapes


def test_square_becomes_a_native_rectangle(tmp_path):
    (slide,) = deck(tmp_path, [kf(square("a", r=1.0))])
    el = find(slide, "!!a")
    assert preset(el) == "rect"
    box = xfrm(el)
    side = 2 * UNIT
    assert box["w"] == pytest.approx(side, abs=2) and box["h"] == pytest.approx(side, abs=2)
    assert box["x"] == pytest.approx((SLIDE_W - side) / 2, abs=2)
    assert box["y"] == pytest.approx((SLIDE_H - side) / 2, abs=2)
    # manim stroke_width 4 -> 0.04 scene units
    assert el.xpath("./p:spPr/a:ln/@w") == [str(round(0.04 * UNIT))]


def test_freeform_mode_writes_manims_cubic_beziers(tmp_path):
    (slide,) = deck(tmp_path, [kf(square("a"))], native_shapes=False)
    el = find(slide, "!!a")
    assert preset(el) is None
    assert len(el.xpath(".//a:cubicBezTo")) == 4
    assert len(el.xpath(".//a:close")) == 1


def test_irregular_outline_stays_a_freeform(tmp_path):
    (slide,) = deck(tmp_path, [kf(blob("b"))])
    el = find(slide, "!!b")
    assert preset(el) is None and len(el.xpath(".//a:cubicBezTo")) == 5


def test_line_becomes_a_connector(tmp_path):
    line = polygon("l", [(-2, 0), (2, 0)], closed=False, fill=None)
    (slide,) = deck(tmp_path, [kf(line)])
    el = find(slide, "!!l")
    assert el.tag == qn("p:cxnSp") and preset(el) == "line"
    assert xfrm(el)["h"] == 0 and xfrm(el)["w"] == pytest.approx(4 * UNIT, abs=2)


def test_translucent_fill_emits_alpha(tmp_path):
    (slide,) = deck(tmp_path, [kf(square("a", fill=(0, 0, 1, 0.5)))])
    assert find(slide, "!!a").xpath("./p:spPr/a:solidFill//a:alpha/@val") == ["50000"]


def test_gradient_fill(tmp_path):
    s = square("a")
    s.fill_gradient = Gradient(((1, 0, 0, 1), (0, 0, 1, 1)), (1.0, 0.0))
    (slide,) = deck(tmp_path, [kf(s)])
    el = find(slide, "!!a")
    assert el.xpath("./p:spPr/a:gradFill/a:gsLst/a:gs/@pos") == ["0", "100000"]
    assert el.xpath("./p:spPr/a:gradFill/a:lin/@ang") == ["0"]


def test_invisible_shape_is_left_off(tmp_path):
    s = square("a", fill=(1, 0, 0, 0), stroke=(1, 1, 1, 0))
    (slide,) = deck(tmp_path, [kf(s)])
    assert top_names(slide) == []


def test_picture_keeps_its_rotation(tmp_path):
    buf = io.BytesIO()
    Image.new("RGB", (4, 2), "red").save(buf, format="PNG")
    pic = ImageShape(
        "img", buf.getvalue(), center=(1.0, 0.0), width=4.0, height=2.0, angle=np.pi / 6
    )
    (slide,) = deck(tmp_path, [kf(pic)])
    shape = slide.shapes[0]
    assert shape.name == "!!img" and shape.rotation == pytest.approx(330)
    assert shape.width == pytest.approx(4 * UNIT, abs=2)


# ---------------------------------------------------------- native motion


def test_rigid_turn_becomes_a_real_rotation(tmp_path):
    slides = deck(tmp_path, [kf(square("a")), kf(square("a", angle=np.radians(-30)))])
    assert len(slides) == 2
    a, b = (find(s, "!!a") for s in slides)
    assert preset(a) == preset(b) == "rect"
    # manim turns counter-clockwise by -30 deg = clockwise 30 on the slide
    assert xfrm(a)["rot"] == 0 and xfrm(b)["rot"] == pytest.approx(30)
    assert xfrm(a)["w"] == pytest.approx(xfrm(b)["w"], abs=2)


def test_rotation_accumulates_past_a_full_turn(tmp_path):
    frames = [kf(blob("b", cx=-1, cy=-1))]
    for step in range(1, 10):
        turned = blob("b", cx=-1, cy=-1)
        c, s = np.cos(step * np.pi / 4), np.sin(step * np.pi / 4)
        turned.subpaths = [sp @ np.array([[c, s], [-s, c]]) for sp in turned.subpaths]
        frames.append(kf(turned))
    slides = deck(tmp_path, frames)
    rots = [xfrm(find(s, "!!b"))["rot"] for s in slides]
    # counter-clockwise in manim is a negative (clockwise-positive) rotation in PowerPoint
    assert rots == pytest.approx([0, 315, 270, 225, 180, 135, 90, 45, 0, 315])


def test_preset_survives_until_the_outline_changes(tmp_path):
    slides = deck(
        tmp_path,
        [kf(circle("a", cx=-3)), kf(circle("a")), kf(square("a")), kf(square("a", cx=3))],
    )
    forms = [(preset(find(s, "!!a")), is_morph(s)) for s in slides]
    assert forms == [
        ("ellipse", False),
        ("ellipse", True),  # a native oval while it only moves
        (None, False),  # instant cut to a freeform with manim's points...
        (None, True),  # ...so that Morph can bend it into the square
        ("rect", False),  # and an instant cut to a native rectangle
        ("rect", True),
    ]
    # the copies advance on their own; the stops still wait for a click
    assert 'advTm="0"' in transition(slides[2]) and 'advTm="0"' in transition(slides[4])
    assert "advTm" not in transition(slides[1]) and "advTm" not in transition(slides[3])


def test_stretching_a_preset_keeps_it_native(tmp_path):
    slides = deck(tmp_path, [kf(circle("a")), kf(circle("a", rx=3.0))])
    assert len(slides) == 2
    assert [preset(find(s, "!!a")) for s in slides] == ["ellipse", "ellipse"]


# ---------------------------------------------------------------- groups


def word(cx=0.0, scale=1.0, gap=1.0):
    parts = []
    for n in range(3):
        b = blob(f"Text 1.{n + 1}", cx=cx + n * 2.5 * scale * gap, scale=scale)
        b.group = "Text 1"
        parts.append(b)
    return parts


def test_compound_moving_as_a_whole_is_one_group(tmp_path):
    slides = deck(tmp_path, [kf(*word()), kf(*word(cx=-3, scale=0.5))])
    for slide in slides:
        assert top_names(slide) == ["!!Text 1"]
        group = find(slide, "!!Text 1")
        assert group.tag == qn("p:grpSp")
        assert group.xpath("./p:sp/p:nvSpPr/p:cNvPr/@name") == ["Text 1.1", "Text 1.2", "Text 1.3"]
    a, b = (xfrm(find(s, "!!Text 1")) for s in slides)
    assert b["w"] == pytest.approx(a["w"] / 2, rel=1e-3)


def test_compound_whose_parts_move_apart_stays_loose(tmp_path):
    slides = deck(tmp_path, [kf(*word()), kf(*word(gap=1.6))])
    assert len(slides) == 2
    for slide in slides:
        assert top_names(slide) == ["!!Text 1.1", "!!Text 1.2", "!!Text 1.3"]


def test_group_then_loose_parts_are_joined_by_a_cut(tmp_path):
    slides = deck(tmp_path, [kf(*word(cx=-2)), kf(*word()), kf(*word(gap=1.6))])
    assert [top_names(s)[0] for s in slides] == ["!!Text 1", "!!Text 1", "!!Text 1.1", "!!Text 1.1"]
    assert [is_morph(s) for s in slides] == [False, True, False, True]


def test_rotating_group_gets_a_rotation(tmp_path):
    turned = word()
    c, s = np.cos(0.5), np.sin(0.5)
    for part in turned:
        part.subpaths = [sp @ np.array([[c, s], [-s, c]]) for sp in part.subpaths]
    slides = deck(tmp_path, [kf(*word()), kf(*turned)])
    assert xfrm(find(slides[1], "!!Text 1"))["rot"] == pytest.approx(
        360 - np.degrees(0.5), abs=1e-3
    )
    # the children do not change inside the group
    before = [xfrm(el) for el in find(slides[0], "!!Text 1").xpath("./p:sp")]
    after = [xfrm(el) for el in find(slides[1], "!!Text 1").xpath("./p:sp")]
    for x, y in zip(before, after, strict=True):
        assert x["w"] == pytest.approx(y["w"], abs=3) and x["rot"] == pytest.approx(
            y["rot"], abs=1e-3
        )


def test_lone_freeform_is_wrapped_only_when_it_changes_size(tmp_path):
    def lone(**kwargs):
        b = blob("Star 1", **kwargs)
        b.group = "Star 1"
        return b

    moved = deck(tmp_path, [kf(lone()), kf(lone(cx=3))])
    assert all(find(s, "!!Star 1").tag == qn("p:sp") for s in moved)
    scaled = deck(tmp_path, [kf(lone()), kf(lone(scale=2))])
    assert all(find(s, "!!Star 1").tag == qn("p:grpSp") for s in scaled)


def test_grouping_can_be_turned_off(tmp_path):
    slides = deck(tmp_path, [kf(*word()), kf(*word(cx=-3))], group=False)
    assert top_names(slides[0]) == ["!!Text 1.1", "!!Text 1.2", "!!Text 1.3"]


# ---------------------------------------------------------------- naming


def test_replacement_inherits_name_when_geometry_matches():
    kfs = [kf(square("old", r=1)), kf(square("new", r=1)), kf(square("new", r=2))]
    assert resolve_names(kfs, FW) == [["old"], ["old"], ["old"]]


def test_unrelated_new_shape_keeps_its_own_name():
    kfs = [kf(square("old", cx=-3)), kf(square("new", cx=3))]
    assert resolve_names(kfs, FW) == [["old"], ["new"]]


# ------------------------------------------------- entrances and exits


def point(key, x=1.0, y=1.0):
    return PathShape(
        key=key,
        subpaths=[np.zeros((4, 2)) + np.array([x, y])],
        closed=[False],
        fill=(1, 0, 0, 1),
        stroke=None,
        stroke_width=0,
    )


def test_collapsed_shape_is_seeded_with_a_transparent_miniature():
    kfs = [kf(point("a")), kf(square("a", cx=1, cy=1, r=1))]
    seeded = prepare_shapes(kfs, resolve_names(kfs, FW), FW)[0][0]
    x0, _, x1, _ = seeded.bounds()
    assert 0 < x1 - x0 < 0.1 and (x0 + x1) / 2 == pytest.approx(1)
    assert len(seeded.subpaths[0]) == 16  # same structure as the target
    assert not seeded.visible and seeded.fill[:3] == (1, 0, 0)


def test_drawn_shape_fades_instead_of_growing_from_a_dot():
    kfs = [kf(point("a"), fade_keys=frozenset({"a"})), kf(square("a", cx=1, cy=1))]
    assert prepare_shapes(kfs, resolve_names(kfs, FW), FW)[0] == [None]


def test_drawn_line_still_grows_from_its_start():
    line = polygon("l", [(1, 1), (4, 1)], closed=False, fill=None)
    kfs = [kf(point("l"), fade_keys=frozenset({"l"})), kf(line)]
    seeded = prepare_shapes(kfs, resolve_names(kfs, FW), FW)[0][0]
    assert seeded is not None and not seeded.visible


def test_transparent_shape_is_kept_only_if_it_moves_into_view():
    hidden = {"fill": (1, 0, 0, 0), "stroke": (1, 1, 1, 0)}
    moving = [kf(square("a", cy=-1, **hidden)), kf(square("a"))]
    assert prepare_shapes(moving, resolve_names(moving, FW), FW)[0][0] is not None
    in_place = [kf(square("a", **hidden)), kf(square("a"))]
    assert prepare_shapes(in_place, resolve_names(in_place, FW), FW)[0] == [None]


def test_departure_state_shares_the_slide_when_nothing_visible_differs(tmp_path):
    arrive = Snapshot([square("a")])
    depart = Snapshot([square("a"), point("b", 3, 0)])  # b is about to grow in
    slides = deck(
        tmp_path, [Keyframe(arrive, departure=depart), kf(square("a"), circle("b", cx=3))]
    )
    assert len(slides) == 2
    assert top_names(slides[0]) == ["!!a", "!!b"]
    assert find(slides[0], "!!b").xpath(".//a:alpha/@val")[0] == "0"


# ------------------------------------------------------------- slide flow


def _tl(mode):
    k = [Keyframe(Snapshot([])) for _ in range(4)]
    k[1].starts_play = True  # play 1 -> slides 1
    k[2].starts_play = False  # play 1 continues (sub-step)
    k[3].starts_play = True  # play 2
    k[2].hold_ms = 700
    k[2].marker = mode == "markers"
    return Timeline(keyframes=k, mode=mode)


def test_click_mode_advance_flags():
    tl = _tl("click")
    assert [advance_after(tl, i) for i in range(4)] == [None, 0, None, None]


def test_marker_mode_advance_flags():
    tl = _tl("markers")
    assert [advance_after(tl, i) for i in range(4)] == [0, 0, None, None]


def test_write_pptx_roundtrip(tmp_path):
    kfs = [
        Keyframe(Snapshot([square("a", cx=-2)], background=(0, 0, 0))),
        Keyframe(Snapshot([square("a", cx=2)]), transition_ms=800, starts_play=True, notes="hi"),
    ]
    out = write_pptx(Timeline(keyframes=kfs, title="Demo"), tmp_path / "t.pptx")
    prs = Presentation(out)
    assert len(prs.slides) == 2
    assert [s.name for s in prs.slides[1].shapes] == ["!!a"]
    xml = prs.slides[1]._element.xml
    assert 'p14:dur="800"' in xml and "p159:morph" in xml
    assert prs.slides[1].notes_slide.notes_text_frame.text == "hi"
    assert prs.core_properties.title == "Demo" and prs.core_properties.author == ""


# ------------------------------------------------------------- the deck


def test_fresh_deck_is_a_proper_dark_16_9_deck(tmp_path):
    out = write_pptx(Timeline(keyframes=[kf(square("a"))]), tmp_path / "t.pptx")
    prs = Presentation(out)
    assert (prs.slide_width, prs.slide_height) == (SLIDE_W, SLIDE_H)
    assert prs.element.find(qn("p:sldSz")).get("type") is None
    master = prs.slide_masters[0]._element
    assert master.xpath("./p:cSld/p:bg//a:srgbClr/@val") == ["000000"]
    clr_map = master.find(qn("p:clrMap"))
    assert (clr_map.get("bg1"), clr_map.get("tx1")) == ("dk1", "lt1")  # light text on dark
    assert not prs.slides[0]._element.xpath("./p:cSld/p:bg")  # the slide inherits it
    # layouts were stretched from 4:3, so a slide added by hand fills the width
    title = next(s for s in prs.slide_masters[0].shapes if s.is_placeholder)
    assert title.left + title.width == pytest.approx(SLIDE_W - title.left, rel=0.01)


def test_odd_background_is_painted_on_its_own_slide(tmp_path):
    a, b = kf(square("a")), kf(square("a", cx=1))
    b.snapshot.background = (1.0, 1.0, 1.0)
    c = kf(square("a", cx=2))
    slides = deck(tmp_path, [a, b, c])
    painted = [bool(s._element.xpath("./p:cSld/p:bg")) for s in slides]
    assert painted == [False, True, False]


def test_template_supplies_size_theme_and_existing_slides(tmp_path):
    template = Presentation()
    template.slide_width, template.slide_height = Emu(9_144_000), Emu(6_858_000)  # 4:3
    template.slides.add_slide(template.slide_layouts[0])
    template.save(tmp_path / "talk.pptx")

    out = write_pptx(
        Timeline(keyframes=[kf(square("a", r=1.0))]),
        tmp_path / "t.pptx",
        template=tmp_path / "talk.pptx",
    )
    prs = Presentation(out)
    assert len(prs.slides) == 2 and prs.slide_width == 9_144_000
    slide = prs.slides[1]
    assert not slide._element.xpath("./p:cSld/p:bg")  # the template's background shows
    box = xfrm(find(slide, "!!a"))
    scale = 9_144_000 / FW  # the 16:9 frame is letterboxed into the 4:3 slide
    assert box["w"] == pytest.approx(2 * scale, abs=2)
    assert box["y"] + box["h"] / 2 == pytest.approx(6_858_000 / 2, abs=2)


def test_sections_start_at_named_markers(tmp_path):
    kfs = [kf(square("a")), kf(square("a", cx=1), section="Proof"), kf(square("a", cx=2))]
    for k in kfs[1:]:
        k.transition_ms = 500
    out = write_pptx(Timeline(keyframes=kfs, title="Talk"), tmp_path / "t.pptx")
    prs = Presentation(out)
    p14 = "{http://schemas.microsoft.com/office/powerpoint/2010/main}"
    sections = list(prs.element.iter(p14 + "section"))
    assert [s.get("name") for s in sections] == ["Talk", "Proof"]
    assert [len(list(s.iter(p14 + "sldId"))) for s in sections] == [1, 2]
