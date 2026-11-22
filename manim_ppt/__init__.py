"""manim-ppt: export Manim animations as native, editable PowerPoint slides."""

from .model import ImageShape, Keyframe, PathShape, Snapshot, Timeline
from .scene import PPTScene, default_samples
from .writer import write_pptx

__all__ = [
    "PPTScene",
    "default_samples",
    "write_pptx",
    "Timeline",
    "Keyframe",
    "Snapshot",
    "PathShape",
    "ImageShape",
]
__version__ = "0.1.0"
