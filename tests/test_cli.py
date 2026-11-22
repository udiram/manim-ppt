"""The ``manim-ppt`` command: plain scenes in, decks out."""

import textwrap

import pytest
from pptx import Presentation

from manim_ppt.cli import main

SCENES = textwrap.dedent(
    """
    from manim import *

    class First(Scene):  # a plain manim scene: no manim-ppt import at all
        def construct(self):
            c = Circle()
            self.play(Create(c))
            self.play(c.animate.shift(RIGHT))

    class Second(Scene):
        def construct(self):
            self.add(Square())
    """
)


@pytest.fixture
def scene_file(tmp_path):
    path = tmp_path / "scenes.py"
    path.write_text(SCENES)
    return path


def test_exports_a_plain_scene(scene_file, tmp_path, capsys):
    out = tmp_path / "out.pptx"
    assert main([str(scene_file), "First", "-o", str(out)]) == 0
    assert len(Presentation(out).slides) == 3
    assert "First: 3 slides" in capsys.readouterr().out
    assert (
        not (tmp_path / "media").exists() and not (scene_file.parent / "media" / "videos").exists()
    )


def test_all_scenes_go_to_a_directory(scene_file, tmp_path):
    assert main([str(scene_file), "--all", "-o", str(tmp_path / "decks")]) == 0
    assert sorted(p.name for p in (tmp_path / "decks").iterdir()) == ["First.pptx", "Second.pptx"]


def test_options_override_scene_attributes(scene_file, tmp_path):
    out = tmp_path / "out.pptx"
    main([str(scene_file), "First", "-o", str(out), "--mode", "markers", "--freeform"])
    slides = list(Presentation(out).slides)
    assert 'advTm="0"' in slides[1]._element.xml  # markers: plays chain without clicks
    assert "prstGeom" not in slides[1]._element.xml


def test_ambiguous_scene_is_an_error(scene_file, capsys):
    with pytest.raises(SystemExit):
        main([str(scene_file)])
    assert "First, Second" in capsys.readouterr().err


def test_unknown_scene_is_an_error(scene_file, capsys):
    with pytest.raises(SystemExit):
        main([str(scene_file), "Nope"])
    assert "no scene named Nope" in capsys.readouterr().err
