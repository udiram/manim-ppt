"""End-to-end: run real manim scenes and inspect the generated deck."""

import pytest
from manim import (
    BLUE,
    DOWN,
    PI,
    RIGHT,
    UP,
    Circle,
    Create,
    Dot,
    FadeIn,
    FadeOut,
    GrowFromCenter,
    Line,
    MovingCameraScene,
    ReplacementTransform,
    Rotate,
    Square,
    Star,
    Text,
    Transform,
    Triangle,
    ValueTracker,
    VGroup,
    Write,
    always_redraw,
    tempconfig,
)
from pptx import Presentation
from pptx.oxml.ns import qn

from manim_ppt import PPTMixin, PPTScene

from .test_writer import find, is_morph, preset, top_names, transition, xfrm


def render(scene_cls, tmp_path):
    """Run a scene without encoding video and return its slides."""
    with tempconfig(
        {
            "dry_run": True,
            "media_dir": str(tmp_path),
            "disable_caching": True,
            "progress_bar": "none",
            "verbosity": "WARNING",
        }
    ):
        scene = scene_cls()
        scene.skip_animation_preview = True
        scene.render()
    return list(Presentation(tmp_path / "pptx" / f"{scene_cls.__name__}.pptx").slides)


def stops(slides):
    """Slides the deck actually rests on (not the instant cuts in between)."""
    return [s for i, s in enumerate(slides) if i == 0 or is_morph(s)]


def test_shapes_get_readable_names_and_native_forms(tmp_path):
    class Names(PPTScene):
        def construct(self):
            self.add(
                Circle().shift(UP), Circle().shift(DOWN), Square(), Line(UP, RIGHT), Text("hi")
            )

    (slide,) = render(Names, tmp_path)
    assert top_names(slide) == ["!!Circle 1", "!!Circle 2", "!!Square 1", "!!Line 1", "!!Text 1"]
    assert preset(find(slide, "!!Circle 1")) == "ellipse"
    assert preset(find(slide, "!!Square 1")) == "rect"
    assert find(slide, "!!Line 1").tag == qn("p:cxnSp")
    text = find(slide, "!!Text 1")
    assert text.tag == qn("p:grpSp")
    assert text.xpath("./p:sp/p:nvSpPr/p:cNvPr/@name") == ["Text 1.1", "Text 1.2"]


def test_transform_keeps_one_morphing_shape(tmp_path):
    class T(PPTScene):
        def construct(self):
            c = Circle(color=BLUE, fill_opacity=0.5)
            self.add(c)
            self.play(c.animate.shift(RIGHT))
            self.play(Transform(c, Square()))
            self.play(c.animate.shift(RIGHT))

    slides = render(T, tmp_path)
    assert all(top_names(s) == ["!!Circle 1"] for s in slides)
    rest = stops(slides)
    assert len(rest) == 4
    # a native oval, bent into a square as freeforms, then a native rectangle
    assert preset(find(rest[0], "!!Circle 1")) == preset(find(rest[1], "!!Circle 1")) == "ellipse"
    assert preset(find(rest[2], "!!Circle 1")) is None
    assert preset(find(rest[3], "!!Circle 1")) == "rect"
    # click mode: every play waits for a click
    assert all("advTm" not in transition(s) for s in rest)


def test_replacement_transform_inherits_name(tmp_path):
    class R(PPTScene):
        def construct(self):
            a = Square(fill_opacity=1)
            self.play(FadeIn(a))
            self.play(ReplacementTransform(a, Triangle(fill_opacity=1)))

    slides = render(R, tmp_path)
    assert top_names(slides[-1]) == top_names(slides[-2]) == ["!!Square 1"]


def test_rotation_is_sampled_and_native(tmp_path):
    class Rot(PPTScene):
        def construct(self):
            s = Square()
            self.add(s)
            self.play(Rotate(s, angle=PI))

    slides = render(Rot, tmp_path)
    # 180 degrees -> 4 Morph steps of 45 degrees, chained automatically.
    assert len(slides) == 5
    assert all('advTm="0"' in transition(s) for s in slides[1:4])
    assert all(preset(find(s, "!!Square 1")) == "rect" for s in slides)
    rots = [xfrm(find(s, "!!Square 1"))["rot"] for s in slides]
    assert rots == pytest.approx([0, 315, 270, 225, 180], abs=6)  # to the nearest frame
    # manim eases in and out, so the outer steps take longer than the inner ones
    durations = [int(transition(s).split('p14:dur="')[1].split('"')[0]) for s in slides[1:]]
    assert durations[0] > durations[1] and durations[3] > durations[2]
    assert sum(durations) == pytest.approx(1000, abs=40)


def test_markers_and_waits(tmp_path):
    class M(PPTScene):
        def construct(self):
            self.next_slide(notes="first")
            c = Circle()
            self.play(GrowFromCenter(c))
            self.wait(0.5)
            self.play(c.animate.shift(RIGHT))
            self.next_slide(notes="second", section="Part two")
            self.play(c.animate.scale(0.5))

    s = render(M, tmp_path)
    assert len(s) == 4
    assert "advTm" not in transition(s[0])  # initial next_slide -> wait
    assert 'advTm="500"' in transition(s[1])  # wait(0.5) becomes a hold
    assert "advTm" not in transition(s[2])  # marker
    assert s[3].notes_slide.notes_text_frame.text == "second"
    p14 = "{http://schemas.microsoft.com/office/powerpoint/2010/main}"
    sections = list(s[0].part.package.presentation_part.presentation.element.iter(p14 + "section"))
    assert [x.get("name") for x in sections] == ["M", "Part two"]


def test_ppt_samples_override(tmp_path):
    class Override(PPTScene):
        def construct(self):
            c = Circle()
            self.add(c)
            self.play(c.animate.shift(RIGHT), ppt_samples=3)

    assert len(render(Override, tmp_path)) == 4


@pytest.mark.parametrize("mode", ["click", "markers"])
def test_forced_mode(tmp_path, mode):
    class F(PPTScene):
        ppt_mode = mode

        def construct(self):
            c = Circle()
            self.add(c)
            self.play(c.animate.shift(RIGHT))
            self.play(c.animate.shift(RIGHT))

    s = render(F, tmp_path)
    has_adv = 'advTm="0"' in transition(s[1])
    assert has_adv == (mode == "markers")


def test_bad_mode_is_rejected():
    class Bad(PPTScene):
        ppt_mode = "sideways"

    with pytest.raises(ValueError, match="ppt_mode"):
        Bad()


def test_written_text_fades_in_as_one_group(tmp_path):
    class W(PPTScene):
        def construct(self):
            t = Text("morph")
            self.play(Write(t))
            self.play(t.animate.scale(0.5).to_edge(UP))

    slides = render(W, tmp_path)
    assert len(slides) == 3
    assert top_names(slides[0]) == []  # nothing yet: Morph fades the text in
    assert top_names(slides[1]) == top_names(slides[2]) == ["!!Text 1"]
    big, small = (xfrm(find(s, "!!Text 1")) for s in slides[1:])
    assert small["w"] == pytest.approx(big["w"] / 2, rel=1e-3)


def test_grow_starts_from_an_invisible_miniature(tmp_path):
    class G(PPTScene):
        def construct(self):
            self.play(GrowFromCenter(Circle(fill_opacity=1)))

    first, last = render(G, tmp_path)
    seed, full = find(first, "!!Circle 1"), find(last, "!!Circle 1")
    assert preset(seed) == preset(full) == "ellipse"
    assert xfrm(seed)["w"] < xfrm(full)["w"] / 20
    assert set(seed.xpath(".//a:alpha/@val")) == {"0"}


def test_created_line_grows_but_created_shape_fades(tmp_path):
    class C(PPTScene):
        def construct(self):
            self.play(Create(Line(UP, RIGHT)), Create(Star()))

    first, last = render(C, tmp_path)
    assert top_names(first) == ["!!Line 1"]
    assert top_names(last) == ["!!Line 1", "!!Star 1"]


def test_shifted_fades_leave_a_ghost_to_morph_from_and_to(tmp_path):
    class Fade(PPTScene):
        def construct(self):
            d = Dot()
            self.play(FadeIn(d, shift=UP))
            self.play(FadeOut(d, shift=RIGHT))

    slides = render(Fade, tmp_path)
    assert len(slides) == 3
    start, shown, gone = (find(s, "!!Dot 1") for s in slides)
    assert start.xpath(".//a:alpha/@val") == ["0"] and gone.xpath(".//a:alpha/@val") == ["0"]
    assert not shown.xpath(".//a:alpha/@val")
    assert xfrm(start)["y"] > xfrm(shown)["y"]  # it rises into place...
    assert xfrm(gone)["x"] > xfrm(shown)["x"]  # ...and leaves to the right


def test_redrawn_compound_keeps_its_names(tmp_path):
    class Redraw(PPTScene):
        def construct(self):
            t = ValueTracker(0)
            pair = always_redraw(
                lambda: VGroup(Dot(RIGHT * t.get_value()), Square().shift(UP * t.get_value()))
            )
            self.add(pair)
            self.play(t.animate.set_value(2), run_time=1)

    slides = render(Redraw, tmp_path)
    assert len(slides) > 2
    assert all(top_names(s) == top_names(slides[0]) for s in slides)


def test_moving_camera_zoom_is_captured(tmp_path):
    class Zoom(PPTMixin, MovingCameraScene):
        def construct(self):
            self.add(Square())
            self.play(self.camera.frame.animate.scale(0.5))

    before, after = render(Zoom, tmp_path)
    a, b = xfrm(find(before, "!!Square 1")), xfrm(find(after, "!!Square 1"))
    assert b["w"] == pytest.approx(2 * a["w"], rel=1e-3)
    assert int(find(after, "!!Square 1").xpath("./p:spPr/a:ln/@w")[0]) == pytest.approx(
        2 * int(find(before, "!!Square 1").xpath("./p:spPr/a:ln/@w")[0]), abs=2
    )


def test_export_options_reach_the_writer(tmp_path):
    class Plain(PPTScene):
        ppt_native_shapes = False
        ppt_group = False
        ppt_output = str(tmp_path / "custom" / "deck.pptx")

        def construct(self):
            self.add(Square(), Text("ab"))

    with tempconfig({"dry_run": True, "media_dir": str(tmp_path), "verbosity": "WARNING"}):
        scene = Plain()
        scene.render()
    assert scene.ppt_path == tmp_path / "custom" / "deck.pptx"
    (slide,) = Presentation(scene.ppt_path).slides
    assert preset(find(slide, "!!Square 1")) is None
    assert top_names(slide) == ["!!Square 1", "!!Text 1.1", "!!Text 1.2"]


def test_a_bit_of_everything_exports(tmp_path):
    """No assertion about looks: every animation family must survive export."""
    from manim import (
        DOWN,
        LEFT,
        ORIGIN,
        AddTextLetterByLetter,
        ApplyWave,
        Arc,
        Brace,
        Circumscribe,
        DrawBorderThenFill,
        Flash,
        FocusOn,
        Indicate,
        MoveAlongPath,
        NumberPlane,
        ShrinkToCenter,
        SpinInFromNothing,
        TransformMatchingShapes,
        Uncreate,
        Unwrite,
        Wiggle,
    )

    class Everything(PPTScene):
        def construct(self):
            t = Text("Hello").to_edge(UP)
            self.play(AddTextLetterByLetter(t))
            sq, ci = Square().shift(LEFT * 3), Circle().shift(RIGHT * 3)
            self.play(DrawBorderThenFill(sq), SpinInFromNothing(ci))
            self.play(Indicate(sq), Wiggle(ci))
            self.play(Circumscribe(t), Flash(ORIGIN))
            self.play(MoveAlongPath(sq, Arc(radius=3, angle=PI)), run_time=1)
            self.play(TransformMatchingShapes(t, Text("Hole").to_edge(UP)))
            v = ValueTracker(0)
            num = always_redraw(lambda: Text(f"{v.get_value():.0f}").shift(DOWN * 2))
            self.add(num)
            self.play(v.animate.set_value(42), run_time=1)
            num.clear_updaters()
            brace, plane = Brace(sq, DOWN), NumberPlane().set_opacity(0.3)
            self.play(FadeIn(plane), GrowFromCenter(brace))
            self.play(ApplyWave(sq), FocusOn(ci))
            self.play(ShrinkToCenter(ci), Uncreate(sq), Unwrite(num), FadeOut(plane, brace))
            self.wait()

    slides = render(Everything, tmp_path)
    assert len(slides) > 20
    names = [n for s in slides for n in top_names(s)]
    assert all(n.startswith("!!") for n in names)
    for s in slides:  # names are unique on a slide, or Morph cannot pair them
        assert len(top_names(s)) == len(set(top_names(s)))
