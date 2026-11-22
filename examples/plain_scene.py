"""An ordinary manim scene: it has never heard of manim-ppt.

The command line exports it anyway:

    manim-ppt examples/plain_scene.py

To drop the animation into your own deck, on your own theme:

    manim-ppt examples/plain_scene.py --template my_talk.pptx -o my_talk_animated.pptx
"""

from manim import *


class Pythagoras(Scene):
    def construct(self):
        a, b = 3, 4
        tri = Polygon(ORIGIN, RIGHT * a, RIGHT * a + UP * b, color=WHITE).move_to(ORIGIN)
        p, _, r = tri.get_vertices()
        sq_a = Square(a, color=BLUE, fill_opacity=0.5).next_to(tri, DOWN, buff=0)
        sq_b = Square(b, color=GREEN, fill_opacity=0.5).next_to(tri, RIGHT, buff=0)
        hyp = r - p
        sq_c = Square(np.linalg.norm(hyp), color=RED, fill_opacity=0.5)
        sq_c.rotate(np.arctan2(hyp[1], hyp[0])).move_to(
            (p + r) / 2 + rotate_vector(hyp, PI / 2) / 2
        )
        figure = VGroup(tri, sq_a, sq_b, sq_c).scale(0.45).shift(LEFT * 2.5)

        self.play(Create(tri))
        self.play(FadeIn(sq_a, shift=UP * 0.3), FadeIn(sq_b, shift=LEFT * 0.3))
        self.play(GrowFromCenter(sq_c))

        labels = (
            VGroup(
                Text("a²", color=BLUE),
                Text("+", color=WHITE),
                Text("b²", color=GREEN),
                Text("=", color=WHITE),
                Text("c²", color=RED),
            )
            .arrange(RIGHT, buff=0.3)
            .scale(1.2)
            .shift(RIGHT * 3.5)
        )
        self.play(LaggedStart(*[FadeIn(part, shift=UP * 0.3) for part in labels], lag_ratio=0.25))
        self.play(Indicate(sq_c), Indicate(labels[-1]))
        self.play(
            figure.animate.shift(RIGHT * 0.5).scale(1.1), labels.animate.scale(0.8).to_edge(UP)
        )
