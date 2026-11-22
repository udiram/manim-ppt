"""End-to-end: render real manim scenes and inspect the generated deck."""

import re

import pytest
from manim import (
    BLUE,
    PI,
    RIGHT,
    Circle,
    FadeIn,
    GrowFromCenter,
    ReplacementTransform,
    Rotate,
    Square,
    Transform,
    Triangle,
    tempconfig,
)
from pptx import Presentation

from manim_ppt import PPTScene


def render(scene_cls, tmp_path):
    with tempconfig(
        {
            "quality": "low_quality",
            "media_dir": str(tmp_path),
            "disable_caching": True,
            "progress_bar": "none",
            "verbosity": "WARNING",
        }
    ):
        scene = scene_cls()
        scene.render()
    return Presentation(tmp_path / "pptx" / f"{scene_cls.__name__}.pptx")


def names(slide):
    return [s.name for s in slide.shapes]


def transition(slide):
    return re.findall(r"<p:transition[^>]*>", slide._element.xml)[0]


def test_transform_keeps_one_morphing_shape(tmp_path):
    class T(PPTScene):
        def construct(self):
            c = Circle(color=BLUE, fill_opacity=0.5)
            self.add(c)
            self.play(Transform(c, Square()))
            self.play(c.animate.shift(RIGHT))

    prs = render(T, tmp_path)
    assert len(prs.slides) == 3
    assert names(prs.slides[0]) == names(prs.slides[1]) == names(prs.slides[2])
    # click mode: every play waits for a click
    assert "advTm" not in transition(prs.slides[0])
    assert "advTm" not in transition(prs.slides[1])


def test_replacement_transform_inherits_name(tmp_path):
    class R(PPTScene):
        def construct(self):
            a = Square(fill_opacity=1)
            self.play(FadeIn(a))
            self.play(ReplacementTransform(a, Triangle(fill_opacity=1)))

    prs = render(R, tmp_path)
    assert names(prs.slides[-1]) == names(prs.slides[-2])


def test_rotation_is_sampled(tmp_path):
    class Rot(PPTScene):
        def construct(self):
            s = Square()
            self.add(s)
            self.play(Rotate(s, angle=PI))

    prs = render(Rot, tmp_path)
    # 180 degrees -> 4 Morph steps of 45 degrees, chained automatically.
    assert len(prs.slides) == 5
    assert all('advTm="0"' in transition(s) for s in list(prs.slides)[1:4])


def test_markers_and_waits(tmp_path):
    class M(PPTScene):
        def construct(self):
            self.next_slide(notes="first")
            c = Circle()
            self.play(GrowFromCenter(c))
            self.wait(0.5)
            self.play(c.animate.shift(RIGHT))
            self.next_slide(notes="second")
            self.play(c.animate.scale(0.5))

    prs = render(M, tmp_path)
    s = list(prs.slides)
    assert len(s) == 4
    assert "advTm" not in transition(s[0])  # initial next_slide -> wait
    assert 'advTm="500"' in transition(s[1])  # wait(0.5) becomes a hold
    assert "advTm" not in transition(s[2])  # marker
    assert s[3].notes_slide.notes_text_frame.text == "second"


def test_ppt_samples_override(tmp_path):
    class O(PPTScene):
        def construct(self):
            c = Circle()
            self.add(c)
            self.play(c.animate.shift(RIGHT), ppt_samples=3)

    assert len(render(O, tmp_path).slides) == 4


@pytest.mark.parametrize("mode", ["click", "markers"])
def test_forced_mode(tmp_path, mode):
    class F(PPTScene):
        ppt_mode = mode

        def construct(self):
            c = Circle()
            self.add(c)
            self.play(c.animate.shift(RIGHT))
            self.play(c.animate.shift(RIGHT))

    s = list(render(F, tmp_path).slides)
    has_adv = 'advTm="0"' in transition(s[1])
    assert has_adv == (mode == "markers")
