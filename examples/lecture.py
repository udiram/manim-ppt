"""Marker mode: one click per next_slide(), waits become timed pauses.

Render with:  manim -ql examples/lecture.py Lecture
"""

from manim import *

from manim_ppt import PPTScene


class Lecture(PPTScene):
    def construct(self):
        self.next_slide(notes="Introduce the function we are going to study.")
        axes = Axes(x_range=[-3, 3, 1], y_range=[-1, 9, 2], x_length=8, y_length=5)
        graph = axes.plot(lambda x: x**2, color=BLUE)
        label = Text("y = x²", font_size=36, color=BLUE).next_to(graph, UR, buff=-0.2)
        self.play(Create(axes))
        self.wait(0.5)
        self.play(Create(graph), FadeIn(label))

        self.next_slide(notes="Slide a point along the curve.")
        t = ValueTracker(-2.5)
        dot = always_redraw(lambda: Dot(axes.c2p(t.get_value(), t.get_value() ** 2), color=YELLOW))
        self.play(FadeIn(dot))
        self.play(t.animate.set_value(2.5), run_time=2)

        self.next_slide(notes="Clear the board.")
        self.play(FadeOut(VGroup(axes, graph, label, dot)))
