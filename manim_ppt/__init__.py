"""manim-ppt: export Manim animations as native, editable PowerPoint slides."""

import warnings

with warnings.catch_warnings():
    # One of manim's dependencies still trips SyntaxWarnings on import; they
    # are noise for someone who just ran `manim-ppt scene.py`.
    warnings.simplefilter("ignore", SyntaxWarning)
    from .model import Gradient, ImageShape, Keyframe, PathShape, Snapshot, Timeline
    from .scene import PPTMixin, PPTScene, default_samples
    from .writer import write_deck, write_pptx

__all__ = [
    "Gradient",
    "ImageShape",
    "Keyframe",
    "PPTMixin",
    "PPTScene",
    "PathShape",
    "Snapshot",
    "Timeline",
    "default_samples",
    "write_deck",
    "write_pptx",
]
__version__ = "0.2.0"
