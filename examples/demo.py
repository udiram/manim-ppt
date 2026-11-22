"""Render with:  manim -ql examples/demo.py Demo
The deck is written to media/pptx/Demo.pptx."""

from manim import *

from manim_ppt import PPTScene


class Demo(PPTScene):
    def construct(self):
        title = Text("manim-ppt", font_size=72)
        self.play(Write(title))
        self.play(title.animate.scale(0.5).to_edge(UP))

        circle = Circle(radius=1.5, color=BLUE, fill_opacity=0.6)
        self.play(GrowFromCenter(circle))

        square = Square(side_length=3, color=ORANGE, fill_opacity=0.6)
        self.play(Transform(circle, square))
        self.play(Rotate(circle, angle=PI / 2))

        tri = Triangle(color=GREEN, fill_opacity=0.6).scale(1.8)
        self.play(ReplacementTransform(circle, tri))

        dots = VGroup(*[Dot(radius=0.15, color=YELLOW) for _ in range(5)]).arrange(RIGHT, buff=0.6)
        dots.next_to(tri, DOWN, buff=0.6)
        self.play(LaggedStart(*[FadeIn(d, shift=UP * 0.5) for d in dots], lag_ratio=0.3))
        self.play(Indicate(tri))
        self.play(FadeOut(tri), FadeOut(dots), FadeOut(title))
