import pytest
import numpy as np

from fairypptx import object_utils
from fairypptx import Application, Shapes, Shape


def test_is_object():
    assert object_utils.is_object("string") is False
    application = Application()
    assert object_utils.is_object(application) is False
    assert object_utils.is_object(application.api) is True


def test_get_type():
    application = Application()
    assert object_utils.get_type(application.api) == "Application"
    with pytest.raises(AttributeError):
        object_utils.get_type("non-Object")


def test_upstream():
    shapes = Shapes()
    App = object_utils.upstream(shapes.api, "Application")
    assert object_utils.get_type(App) == "Application"

    Pres = object_utils.upstream(shapes.api, "presentation")
    assert object_utils.get_type(Pres) == "Presentation"

    with pytest.raises(ValueError):
        object_utils.upstream(shapes.api, "NonObject")

    with pytest.raises(ValueError):
        object_utils.upstream("InvalidObject", "NonObject")

    with pytest.raises(ValueError):
        object_utils.upstream(1000, "NonObject")

def test_stored():
    shape = Shape.make(1)

    # Specify `str`.
    shape.api.Line.Weight = 3
    with object_utils.stored(shape.api, "Line.Weight"):
        shape.api.LineWeight = 5
    assert shape.api.Line.Weight  == 3

    # Specified by `Sequence`.
    shape.api.Line.Weight = 2
    shape.api.Fill.ForeColor.RGB = 1
    with object_utils.stored(shape.api, ("Line.Weight", "Fill.ForeColor.RGB")):
        shape.line = 5
        shape.api.Fill.ForeColor.RGB = 100
    assert shape.api.Line.Weight == 2
    assert shape.api.Fill.ForeColor.RGB == 1

    # Handling of Exceptions.
    shape.api.Line.Weight = 4
    with pytest.raises(ValueError):
        with object_utils.stored(shape.api, "Line.Weight"):
            shape.api.Line.Weight = 20
            raise ValueError
    assert shape.api.Line.Weight == 4
    
def test_setattr():
    # Simple specification
    shape = Shape.make(1)
    object_utils.setattr(shape.api, "Left", 200)
    assert shape.api.Left == 200

    # Dot specification
    object_utils.setattr(shape.api, "TextFrame.TextRange.Text", "SampleText")
    assert shape.api.TextFrame.TextRange.Text == "SampleText"

    # Sequence specification
    object_utils.setattr(shape.api, ["TextFrame", "TextRange", "Text"], "SampleText2")
    assert shape.api.TextFrame.TextRange.Text == "SampleText2"

    # This is the extension of `builtins.setattr`.
    # Hence, this is applicable to Non-Object instances.
    object_utils.setattr(shape, "AnyAttribute", None)
    assert shape.AnyAttribute == None


def test_getattr():
    # Simple.
    shape = Shape.make(1)
    shape.api.Left = 129
    assert object_utils.getattr(shape.api, "Left", 129)

    # Dot Specification
    shape.text = "SampleText"
    assert object_utils.getattr(shape.api, "TextFrame.TextRange.Text") == "SampleText"

    # Sequence Specification
    shape.text = "SampleText"
    assert object_utils.getattr(shape.api, ["TextFrame", "TextRange", "Text"]) == "SampleText"

    # This is the extension of `builtins.getattr`.
    # Hence, this is applicable to Non-Object instances.
    array = np.zeros(1, dtype=int)
    assert object_utils.getattr(array, "dtype") == int

    # If `default` is not set and attribute does not exist, then AttributeError.
    with pytest.raises(AttributeError):
        object_utils.getattr(shape, "NON-Attribute")

    # If `default` is set, it returns `default`
    assert object_utils.getattr(shape, "NON-Attribute", 797) == 797


def test_hasattr():
    shape = Shape.make(1)
    assert object_utils.hasattr(shape.api, "Left") == True
    assert object_utils.hasattr(shape.api, "NONATTRIBUTE") == False
    assert object_utils.hasattr(shape.api, "TextFrame.TextRange.Text") == True
    assert object_utils.hasattr(shape.api, "TextFrame.TextRange.None") == False
    

if __name__ == "__main__":
    pytest.main([__file__, "--capture=no"])
