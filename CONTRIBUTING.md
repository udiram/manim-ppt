# Contributing

Thanks for helping out. Issues with a small scene that reproduces the problem
are the most useful kind.

## Setup

```bash
uv venv && uv pip install -e ".[dev]"
pytest
ruff check . && ruff format --check .
```

(`python -m venv .venv && pip install -e ".[dev]"` works just as well.)

## Layout

| file | job | imports manim? |
|---|---|---|
| `manim_ppt/scene.py` | `PPTMixin`/`PPTScene`: hooks into `play()` and records keyframes | yes |
| `manim_ppt/capture.py` | reads a scene into a `Snapshot`; names mobjects | yes |
| `manim_ppt/model.py` | `PathShape`, `Snapshot`, `Keyframe`, `Timeline` | no |
| `manim_ppt/geometry.py` | rotation tracking, preset recognition | no |
| `manim_ppt/writer.py` | plans the form of every shape and writes the `.pptx` | no |
| `manim_ppt/cli.py` | the `manim-ppt` command | lazily |

Tests in `tests/test_writer.py` and `tests/test_geometry.py` build timelines by
hand, so most behaviour can be tested without rendering anything.

## Checking a change in PowerPoint

The test suite checks the XML. Only PowerPoint can tell you whether Morph
actually does what you hoped, and it has surprises in store:

- Morph pairs objects by name (the `!!` prefix); it does **not** pair the
  children of two groups with each other.
- Presets and groups move, scale and rotate cleanly. Freeforms develop
  artifacts as soon as their size changes.
- Opacity is tweened, so a fully transparent shape is a usable start or end
  state.
- Rotation takes the short way round.

When you change how shapes are written, export `examples/demo.py`, play it in
PowerPoint's slideshow and watch the transitions you touched.

## Releasing

1. Bump `__version__` in `manim_ppt/__init__.py` and add a section to
   `CHANGELOG.md`.
2. Tag and publish a GitHub release. The `release` workflow builds the package
   and uploads it to PyPI through
   [trusted publishing](https://docs.pypi.org/trusted-publishers/), which has
   to be set up once for the `manim-ppt` project on PyPI (workflow
   `release.yml`, environment `pypi`).
