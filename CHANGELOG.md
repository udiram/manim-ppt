# Changelog

## 0.2.0

The output is now built from the objects a PowerPoint user would have drawn, and
it plays far more cleanly.

### Added
- **Native shapes.** Rectangles, rounded rectangles, ellipses, triangles and
  lines are written as PowerPoint presets, and rotation is stored as a real
  rotation instead of being baked into the points.
- **Groups.** A compound mobject (a word, an equation, a set of axes) that moves
  as a whole is one PowerPoint group. Morph scales groups cleanly, where it
  mangles loose glyph outlines.
- **Per-transition forms.** A shape is a preset while it moves and a freeform
  only for the transition that bends it into something else; slides that need
  both forms are written twice with an instant cut in between.
- **`manim-ppt` command** that exports any scene, including plain `Scene`
  subclasses, without rendering video.
- **Templates** (`ppt_template`, `--template`): build on your own deck, theme
  and slide size; existing slides are kept.
- **Sections** via `next_slide(section=...)`.
- `PPTMixin` for `MovingCameraScene`, `ThreeDScene` and other scene types.
  Camera pans and zooms are captured; 3D scenes are captured as projected.
- Readable shape names (`Circle 1`, `Text 2.3`) in the Selection Pane, stable
  even for mobjects rebuilt every frame (`always_redraw`, `DecimalNumber`).
- Gradient fills and strokes, rotated pictures, crisp small images.
- Shifted and scaled fades (`FadeIn(m, shift=UP)`, `FadeOut(m, scale=2)`) and
  exits such as `ShrinkToCenter` now move, by leaving a transparent copy for
  Morph to tween from or to.

### Changed
- `Write`, `DrawBorderThenFill` and `Create` (on anything but a straight line)
  fade in instead of zooming out of dots. No more specks on the slide before
  the click.
- Keyframes inside one `play()` are spaced evenly in animation progress, so a
  half turn in four steps is four 45° steps.
- A fresh deck is a proper 16:9 deck: the background lives on the slide master,
  layouts are sized for 16:9, and on dark backgrounds the theme's text is light.
- `ImageShape` is described by center, size and angle instead of a bounding box.
- `Keyframe` gained `departure`, `section` and `fade_keys`; `PathShape` gained
  `group` and gradients.

## 0.1.0

Initial release: `PPTScene`, freeform shapes, Morph transitions, click and
marker modes, speaker notes.
