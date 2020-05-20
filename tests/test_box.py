import pytest
import numpy as np
from fairypptx.box import Box, Interval, EmptySet, intersection_over_union, intersection_over_cover

def test_box():
    box = Box(10, 11, 12, 13)
    assert box.Left == 10
    assert box.Top == 11
    assert box.Width == 12
    assert box.Height == 13

    box = Box((10, 11, 12, 13))
    assert box.Left == 10
    assert box.Top == 11
    assert box.Width == 12
    assert box.Height == 13

    x_interval = Interval(10, 10 + 12)
    y_interval = Interval(11, 11 + 13)
    box = Box(y_interval, x_interval)
    assert box.x_interval == x_interval

    box1 = Box(1, 2, 3, 4)
    box2 = Box(2, 3, 4, 5)
    expected = Box(2, 3, 2, 3)
    assert expected  == Box.intersection(box1, box2)

    box1 = Box(1, 2, 3, 4)
    box2 = Box(2, 3, 4, 5)
    expected = Box(2, 3, 2, 3)
    assert expected  == Box.intersection(box1, box2)

    box1 = Box(1, 2, 3, 4)
    box2 = Box(2, 3, 4, 5)
    expected = Box(1, 2, 5, 6)
    assert expected  == Box.cover(box1, box2)


def test_interval():
    interval = Interval(0, 5)
    assert interval.length == 5
    assert interval.start == 0
    assert interval.end == 5
    interval = Interval((-1, 3))
    assert interval.length == 4

    interval = Interval((3, 1))
    assert interval.length == 2
    # For reversed args, , automatically, range is reversed. 
    assert interval.start == 1
    assert interval.end == 3

    with pytest.raises(TypeError):
        interval = Interval(13)

    with pytest.raises(TypeError):
        interval = Interval("string")

    # Cover's test.
    interval1 = Interval(0, 2) 
    interval2 = Interval(-1, 1)
    interval = Interval.cover(interval1, interval2)
    assert interval.start == -1
    assert interval.end == 2
    
    interval1 = Interval(0, 2) 
    interval2 = Interval(-1, 1)
    interval = interval1.cover(interval2)
    assert interval.start == -1
    assert interval.end == 2

    interval1 = Interval(0, 2) 
    interval2 = Interval(-1, 1)
    with pytest.raises(EmptySet):
        interval = Interval.cover([])

    # Intersection's test.
    interval1 = Interval(0, 2) 
    interval2 = Interval(-1, 1)
    interval = Interval.intersection(interval1, interval2)
    assert interval.start == 0
    assert interval.end == 1

    interval1 = Interval(3, 2) 
    interval2 = Interval(-1, 1)
    with pytest.raises(EmptySet):
        Interval.intersection(interval1, interval2)

    interval1 = Interval(2, 4) 
    interval2 = Interval(1, 2)
    interval = Interval.intersection(interval1, interval2)
    assert interval.start == interval.end == 2

    interval1 = Interval(2, 4) 
    interval2 = Interval(1, 2)
    assert not interval1.issubset(interval2)

    interval1 = Interval(3, 3.5)
    interval2 = Interval(2, 4) 
    assert interval1.issubset(interval2)
    assert interval2.issuperset(interval)


def test_iou():
    # Disjoint cases.
    d1 = {"Left": 1, "Top": 2, "Width": 3, "Height": 4}
    b1 = Box(d1)
    d2 = {"Left": 10, "Top": 20, "Width": 30, "Height": 40}
    b2 = Box(d2)
    assert intersection_over_union(b1, b2) == 0
    assert intersection_over_union(b1, b2, axis=0) == 0
    assert intersection_over_union(b1, b2, axis=1) == 0

    b1 = Box({"Left": 1, "Top": 2, "Width": 3, "Height": 4})
    b2 = Box({"Left": 3, "Top": 4, "Width": 4, "Height": 7})

    # axis == 0
    assert b1.y_interval == Interval(2, 6)
    assert b2.y_interval == Interval(4, 11)
    assert intersection_over_union(b1, b2, axis=0) == Interval(2, 4).length / Interval(2, 11).length

    # axis == 1
    assert b1.x_interval == Interval(1, 4)
    assert b2.x_interval == Interval(3, 7)
    assert intersection_over_union(b1, b2, axis=1) == Interval(3, 4).length / Interval(1, 7).length

    # axis == None
    nominator = Box(Interval(2, 4), Interval(3, 4)).area
    denominator = b1.area + b2.area - nominator 
    assert intersection_over_union(b1, b2, axis=None) ==  nominator / denominator


def test_ioc():
    # Disjoint cases.
    d1 = {"Left": 1, "Top": 2, "Width": 3, "Height": 4}
    b1 = Box(d1)
    d2 = {"Left": 10, "Top": 20, "Width": 30, "Height": 40}
    b2 = Box(d2)
    assert intersection_over_cover(b1, b2) == 0
    assert intersection_over_cover(b1, b2, axis=0) == 0
    assert intersection_over_cover(b1, b2, axis=1) == 0

    b1 = Box({"Left": 1, "Top": 2, "Width": 3, "Height": 4})
    b2 = Box({"Left": 3, "Top": 4, "Width": 4, "Height": 7})

    # For interval, (when axis is specified),  
    # union and color is the same.
    # (Though, mathematically, this is not correct....)

    # axis == 0
    assert b1.y_interval == Interval(2, 6)
    assert b2.y_interval == Interval(4, 11)
    assert intersection_over_cover(b1, b2, axis=0) == Interval(2, 4).length / Interval(2, 11).length

    # axis == 1
    assert b1.x_interval == Interval(1, 4)
    assert b2.x_interval == Interval(3, 7)
    assert intersection_over_cover(b1, b2, axis=1) == Interval(3, 4).length / Interval(1, 7).length

    # axis == None
    nominator = Box(Interval(2, 4), Interval(3, 4)).area
    denominator = Box(Interval(2, 11), Interval(1, 7)).area
    assert intersection_over_cover(b1, b2, axis=None) ==  nominator / denominator


if __name__ == "__main__":

    from PIL import Image
    pytest.main([__file__, "--capture=no"])
