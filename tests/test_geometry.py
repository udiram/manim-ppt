"""Rotation tracking and preset recognition (slide space: y down, clockwise)."""

import math

import numpy as np
import pytest

from manim_ppt.geometry import (
    classify,
    is_straight_line,
    natural_angle,
    rotation,
    similarity,
    similarity_angle,
    to_local,
)

from .shapes import blob, circle, polygon, square


def pts(shape):
    return shape.subpaths


def kind(shape, theta=None):
    theta = natural_angle(pts(shape), shape.closed) if theta is None else theta
    preset = classify(to_local(pts(shape), theta), shape.closed)
    return preset.kind if preset else None


def test_square_is_a_rect_and_circle_an_ellipse():
    assert kind(square("s")) == "rect"
    assert kind(circle("c")) == "ellipse"
    assert kind(circle("e", r=1, rx=2.5)) == "ellipse"
    assert kind(blob("b")) is None


@pytest.mark.parametrize("deg", [10, 30, 45, -20])
def test_rotated_square_is_a_rotated_rect(deg):
    s = square("s", angle=math.radians(deg))
    theta = natural_angle(pts(s), s.closed)
    assert math.degrees(theta) == pytest.approx(deg, abs=1e-6)
    assert kind(s, theta) == "rect"
    assert kind(s, 0.0) is None  # in an unrotated box it is just a diamond


def test_axis_aligned_square_has_no_rotation():
    s = square("s")
    assert natural_angle(pts(s), s.closed) == 0.0


def test_triangle_finds_its_base_and_apex():
    up = polygon("t", [(0, 0), (2, 0), (1, -1.7)])  # y down: apex at the top
    preset = classify(to_local(pts(up), natural_angle(pts(up), up.closed)), up.closed)
    assert preset.kind == "triangle" and preset.adj == pytest.approx(0.5)
    assert natural_angle(pts(up), up.closed) == pytest.approx(0.0)

    right = polygon("r", [(0, 0), (3, 0), (0, -2)])  # right angle at the origin
    theta = natural_angle(pts(right), right.closed)
    preset = classify(to_local(pts(right), theta), right.closed)
    assert preset.kind == "triangle"


def test_line_is_recognised_at_any_angle():
    line = polygon("l", [(0, 0), (3, 4)], closed=False, fill=None)
    assert is_straight_line(pts(line), line.closed)
    theta = natural_angle(pts(line), line.closed)
    assert math.degrees(theta) == pytest.approx(math.degrees(math.atan2(4, 3)))
    frame = to_local(pts(line), theta)
    assert classify(frame, line.closed).kind == "line"
    assert frame.w == pytest.approx(5.0) and frame.h == pytest.approx(0.0, abs=1e-9)


def test_bent_path_is_not_a_line():
    bent = polygon("l", [(0, 0), (3, 0), (3, 2)], closed=False, fill=None)
    assert not is_straight_line(pts(bent), bent.closed)
    assert kind(bent) is None


def test_rounded_rectangle_reports_its_corner_radius():
    w, h, r = 4.0, 2.0, 0.5
    k = 4 / 3 * math.tan(math.pi / 8)
    centers = [(w - r, r, -90), (w - r, h - r, 0), (r, h - r, 90), (r, r, 180)]
    path, cursor = [], None
    for cx, cy, start in centers:
        a0, a1 = math.radians(start), math.radians(start + 90)
        p0 = np.array([cx + r * math.cos(a0), cy + r * math.sin(a0)])
        p1 = np.array([cx + r * math.cos(a1), cy + r * math.sin(a1)])
        t0 = np.array([-math.sin(a0), math.cos(a0)]) * r * k
        t1 = np.array([-math.sin(a1), math.cos(a1)]) * r * k
        if cursor is not None:
            path += [cursor, cursor + (p0 - cursor) / 3, cursor + 2 * (p0 - cursor) / 3, p0]
        path += [p0, p0 + t0, p1 - t1, p1]
        cursor = p1
    first = path[0]
    path += [cursor, cursor + (first - cursor) / 3, cursor + 2 * (first - cursor) / 3, first]
    preset = classify(to_local([np.array(path)], 0.0), [True])
    assert preset.kind == "roundRect"
    assert preset.adj == pytest.approx(r / min(w, h), abs=0.01)


def test_similarity_recovers_rotation_and_scale():
    a = blob("a")
    turn = math.radians(37)
    moved = [sp @ rotation(turn).T * 1.5 + [3, -2] for sp in pts(a)]
    found = similarity(pts(a), moved)
    assert np.angle(found) == pytest.approx(turn)
    assert abs(found) == pytest.approx(1.5)
    assert similarity_angle(pts(a), moved) == pytest.approx(turn)


def test_similarity_rejects_deformation_and_points():
    a = blob("a")
    stretched = [sp * [2.0, 1.0] for sp in pts(a)]
    assert similarity(pts(a), stretched) is None
    collapsed = [sp * 0 for sp in pts(a)]
    assert similarity(pts(a), collapsed) is None
    assert similarity(pts(a), pts(square("s"))) is None  # different structure


def test_local_frame_round_trips_through_rotation():
    s = square("s", cx=3, cy=1, r=2, angle=0.4)
    frame = to_local(pts(s), 0.4)
    assert frame.w == pytest.approx(4) and frame.h == pytest.approx(4)
    center = np.array([frame.x + frame.w / 2, frame.y + frame.h / 2])
    assert center == pytest.approx([3, 1])
    # DrawingML: rotate the unrotated box about its center.
    world = (frame.subpaths[0] + [frame.x, frame.y] - center) @ rotation(0.4).T + center
    assert world == pytest.approx(pts(s)[0])
