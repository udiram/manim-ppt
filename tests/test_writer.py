"""Tests for the manim-independent half: geometry, naming, slide flow."""

import re

import numpy as np
from pptx import Presentation

from manim_ppt.model import Keyframe, PathShape, Snapshot, Timeline
from manim_ppt.writer import (
    _Coords,
    advance_after,
    path_shape_xml,
    resolve_names,
    seed_degenerate,
    write_pptx,
)

FW, FH = 14.222222222222221, 8.0


def square(key, cx=0.0, cy=0.0, r=1.0, fill=(1, 0, 0, 1)):
    corners = [(cx - r, cy - r), (cx + r, cy - r), (cx + r, cy + r), (cx - r, cy + r)]
    pts = []
    for a, b in zip(corners, corners[1:] + corners[:1]):
        a, b = np.array(a), np.array(b)
        pts += [a, a + (b - a) / 3, a + 2 * (b - a) / 3, b]
    return PathShape(
        key=key, subpaths=[np.array(pts)], closed=[True], fill=fill, stroke=(1, 1, 1, 1), stroke_width=4
    )


def coords():
    return _Coords(FW, FH, 12_192_000, 6_858_000)


def test_path_xml_uses_cubic_beziers_and_emu_placement():
    xml = path_shape_xml(square("a", r=1.0), "!!a", 2, coords())
    assert xml.count("<a:cubicBezTo>") == 4
    assert "<a:close/>" in xml
    # A 2x2-unit square centered on the frame: 2 units = 2 * 12192000/14.222 EMU.
    side = round(2 * 12_192_000 / FW)
    m = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"/>', xml)
    assert abs(int(m.group(1)) - side) <= 2 and abs(int(m.group(2)) - side) <= 2
    m = re.search(r'<a:off x="(\d+)" y="(\d+)"/>', xml)
    assert abs(int(m.group(1)) - (12_192_000 - side) / 2) <= 2
    # manim stroke_width 4 -> 0.04 scene units
    assert f'w="{round(0.04 * 12_192_000 / FW)}"' in xml
    assert 'name="!!a"' in xml


def test_invisible_shape_is_skipped():
    s = square("a", fill=(1, 0, 0, 0))
    s.stroke = (1, 1, 1, 0)
    assert path_shape_xml(s, "!!a", 2, coords()) is None


def test_translucent_fill_emits_alpha():
    xml = path_shape_xml(square("a", fill=(0, 0, 1, 0.5)), "!!a", 2, coords())
    assert '<a:alpha val="50000"/>' in xml


def test_replacement_inherits_name_when_geometry_matches():
    kfs = [
        Keyframe(Snapshot([square("old", r=1)])),
        Keyframe(Snapshot([square("new", r=1)])),  # same place, new mobject
        Keyframe(Snapshot([square("new", r=2)])),
    ]
    names = resolve_names(kfs, FW)
    assert names == [["old"], ["old"], ["old"]]


def test_unrelated_new_shape_keeps_its_own_name():
    kfs = [
        Keyframe(Snapshot([square("old", cx=-3)])),
        Keyframe(Snapshot([square("new", cx=3)])),
    ]
    assert resolve_names(kfs, FW) == [["old"], ["new"]]


def test_degenerate_shape_is_seeded_from_next_state():
    point = PathShape(
        key="a", subpaths=[np.zeros((4, 2)) + [1.0, 1.0]], closed=[False],
        fill=(1, 0, 0, 1), stroke=None, stroke_width=0,
    )
    kfs = [Keyframe(Snapshot([point])), Keyframe(Snapshot([square("a", cx=1, cy=1, r=1)]))]
    names = resolve_names(kfs, FW)
    seeded = seed_degenerate(kfs, names, FW)[0][0]
    x0, y0, x1, y1 = seeded.bounds()
    assert 0 < x1 - x0 < 0.1 and abs((x0 + x1) / 2 - 1) < 1e-6
    assert len(seeded.subpaths[0]) == 16  # same structure as the target


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
    out = write_pptx(Timeline(keyframes=kfs), tmp_path / "t.pptx")
    prs = Presentation(out)
    assert len(prs.slides) == 2
    assert [s.name for s in prs.slides[1].shapes] == ["!!a"]
    xml = prs.slides[1]._element.xml
    assert 'p14:dur="800"' in xml and "p159:morph" in xml
    assert prs.slides[1].notes_slide.notes_text_frame.text == "hi"
