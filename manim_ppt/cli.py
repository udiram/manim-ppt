"""``manim-ppt``: turn the scenes in a file into PowerPoint decks.

Works on any manim scene, including ones that have never heard of manim-ppt::

    manim-ppt scene.py MyScene

No video is rendered, so this is much faster than a normal ``manim`` run.
"""

from __future__ import annotations

import argparse
import importlib.util
import inspect
import subprocess
import sys
from collections.abc import Sequence
from pathlib import Path
from typing import Any


def build_parser() -> argparse.ArgumentParser:
    from . import __version__

    parser = argparse.ArgumentParser(
        prog="manim-ppt",
        description=(
            "Export Manim scenes as native, editable PowerPoint decks animated with Morph. "
            "Any Scene works; it does not have to subclass PPTScene."
        ),
    )
    parser.add_argument("file", type=Path, help="Python file containing the scene(s)")
    parser.add_argument(
        "scenes",
        nargs="*",
        metavar="SCENE",
        help="scene class names (default: the only scene in the file)",
    )
    parser.add_argument("-a", "--all", action="store_true", help="export every scene in the file")
    parser.add_argument(
        "-o", "--output", type=Path, help="output .pptx (one scene) or directory (several)"
    )
    parser.add_argument(
        "-t", "--template", type=Path, help="a .pptx whose theme, size and slides to build on"
    )
    parser.add_argument(
        "--mode",
        choices=["auto", "click", "markers"],
        help="slide flow: click per play(), or per next_slide() marker",
    )
    background = parser.add_mutually_exclusive_group()
    background.add_argument(
        "--background",
        dest="background",
        action="store_true",
        default=None,
        help="paint the scene background onto the slides (default without a template)",
    )
    background.add_argument(
        "--no-background",
        dest="background",
        action="store_false",
        help="leave the background to the template or theme",
    )
    parser.add_argument(
        "--freeform",
        action="store_true",
        help="write every shape as a freeform instead of native presets and rotations",
    )
    parser.add_argument(
        "--no-groups",
        action="store_true",
        help="never group the parts of a compound mobject (a word, an equation...)",
    )
    parser.add_argument(
        "--max-samples",
        type=int,
        metavar="N",
        help="cap on Morph steps for a single play() (default 12)",
    )
    parser.add_argument(
        "--video", action="store_true", help="also render the video, like a normal manim run"
    )
    parser.add_argument(
        "-q",
        "--quality",
        choices=["l", "m", "h", "p", "k"],
        default="l",
        help="video quality when --video is given (default: l)",
    )
    parser.add_argument("--open", action="store_true", help="open the deck when done")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser


def load_scenes(file: Path) -> dict[str, type]:
    """Import ``file`` and return the scene classes it defines, in order."""
    from manim import Scene

    file = file.resolve()
    if not file.is_file():
        raise SystemExit(f"manim-ppt: no such file: {file}")
    sys.path.insert(0, str(file.parent))
    spec = importlib.util.spec_from_file_location(file.stem, file)
    if spec is None or spec.loader is None:
        raise SystemExit(f"manim-ppt: cannot import {file}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return {
        name: cls
        for name, cls in vars(module).items()
        if inspect.isclass(cls) and issubclass(cls, Scene) and cls.__module__ == module.__name__
    }


def as_ppt_scene(cls: type, overrides: dict[str, Any]) -> type:
    """``cls`` with PowerPoint export mixed in (if it lacks it) and the
    command-line options applied on top of its own ``ppt_*`` attributes."""
    from .scene import PPTMixin

    bases = (cls,) if issubclass(cls, PPTMixin) else (PPTMixin, cls)
    return type(cls.__name__, bases, {"__module__": cls.__module__, **overrides})


def export(cls: type, output: Path, *, video: bool = False, quality: str = "l") -> tuple[Path, int]:
    """Run ``cls`` and write its deck to ``output``. Returns (path, slides)."""
    from manim import tempconfig

    qualities = {
        "l": "low_quality",
        "m": "medium_quality",
        "h": "high_quality",
        "p": "production_quality",
        "k": "fourk_quality",
    }
    settings: dict[str, Any] = {"disable_caching": True}
    if video:
        settings["quality"] = qualities[quality]
    else:
        # Nothing is encoded, so the frame rate only decides how exactly the
        # keyframes line up with their requested times.
        settings.update(dry_run=True, frame_rate=60, progress_bar="none", verbosity="WARNING")
    with tempconfig(settings):
        scene = cls()
        scene.ppt_output = output
        if not video:
            scene.skip_animation_preview = True  # step through time without drawing frames
        scene.render()
        return scene.ppt_path, scene.ppt_slide_count


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    found = load_scenes(args.file)
    if not found:
        parser.error(f"{args.file} defines no Scene classes")
    if args.all:
        chosen = list(found)
    elif args.scenes:
        missing = [name for name in args.scenes if name not in found]
        if missing:
            parser.error(
                f"no scene named {', '.join(missing)} in {args.file} (found: {', '.join(found)})"
            )
        chosen = args.scenes
    elif len(found) == 1:
        chosen = list(found)
    else:
        parser.error(
            f"{args.file} has several scenes; name one or pass --all (found: {', '.join(found)})"
        )

    if args.template is not None and not args.template.is_file():
        parser.error(f"template not found: {args.template}")
    overrides: dict[str, Any] = {}
    if args.mode:
        overrides["ppt_mode"] = args.mode
    if args.template:
        overrides["ppt_template"] = args.template
    if args.background is not None:
        overrides["ppt_background"] = args.background
    if args.freeform:
        overrides["ppt_native_shapes"] = False
    if args.no_groups:
        overrides["ppt_group"] = False
    if args.max_samples:
        overrides["ppt_max_samples"] = args.max_samples

    out = args.output
    single_file = out is not None and out.suffix.lower() == ".pptx"
    if single_file and len(chosen) > 1:
        parser.error("--output must be a directory when exporting several scenes")

    last: Path | None = None
    for name in chosen:
        target = out if single_file else (out or Path.cwd()) / f"{name}.pptx"
        path, slides = export(
            as_ppt_scene(found[name], overrides), target, video=args.video, quality=args.quality
        )
        print(f"{name}: {slides} slide{'s' * (slides != 1)} -> {path}")
        last = path

    if args.open and last is not None:
        _open(last)
    return 0


def _open(path: Path) -> None:
    if sys.platform == "darwin":
        subprocess.run(["open", str(path)], check=False)
    elif sys.platform == "win32":
        import os

        os.startfile(str(path))  # type: ignore[attr-defined]
    else:
        subprocess.run(["xdg-open", str(path)], check=False)


if __name__ == "__main__":
    raise SystemExit(main())
