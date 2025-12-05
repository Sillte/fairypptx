import pytest

from fairypptx.shape_range import ShapeRange
from fairypptx.shape import Shape
from fairypptx import object_utils


@pytest.fixture(scope="module")
def two_shapes():
    # Create two shapes on a slide using the real PowerPoint COM API.
    # These tests intentionally avoid mocking and require PowerPoint to be available.
    try:
        s1 = Shape.make_textbox("Alpha")
        s2 = Shape.make_textbox("Beta")
    except Exception as e:
        pytest.skip(f"PowerPoint not available or cannot create shapes: {e}")

    try:
        yield s1, s2
    finally:
        # Cleanup: delete shapes if still present
        try:
            s1.api.Delete()
        except Exception:
            pass
        try:
            s2.api.Delete()
        except Exception:
            pass


def test_shape_range_from_shape_instances(two_shapes):
    s1, s2 = two_shapes
    sr = ShapeRange([s1, s2])
    assert len(sr) == 2
    assert list(sr)[0].api.Id == s1.api.Id
    assert list(sr)[1].api.Id == s2.api.Id


def test_shape_range_from_com_objects(two_shapes):
    s1, s2 = two_shapes
    sr = ShapeRange([s1.api, s2.api])
    assert len(sr) == 2
    assert all(isinstance(elem, Shape) for elem in sr)


def test_getitem_and_slice(two_shapes):
    s1, s2 = two_shapes
    sr = ShapeRange([s1, s2])
    first = sr[0]
    assert isinstance(first, Shape)
    assert first.api.Id == s1.api.Id
    slice_sr = sr[0:2]
    assert isinstance(slice_sr, ShapeRange)
    assert len(slice_sr) == 2


def test_api_property_returns_com_shaperange(two_shapes):
    s1, s2 = two_shapes
    sr = ShapeRange([s1, s2])
    api = sr.api
    assert object_utils.is_object(api, "ShapeRange")


def test_empty_shaperange_api_raises():
    sr = ShapeRange([])
    with pytest.raises(ValueError):
        _ = sr.api


def test_leafs():
    from fairypptx import Slides
    from fairypptx import constants
    Slides().add(layout=constants.ppLayoutBlank)
    s1 = Shape.make("S1")
    s2 = Shape.make("S2")
    sg = ShapeRange([s1, s2])
    grouped = sg.group()
    shapes = ShapeRange([grouped]).leafs
    assert len(shapes) == 2
    texts = {str(shape.text) for shape in shapes}
    assert texts == {"S1", "S2"}
