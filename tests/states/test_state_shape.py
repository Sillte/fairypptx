from fairypptx.shape import Shape, TableShape, GroupShape
from fairypptx.shapes import Shapes
from fairypptx.slide import Slide
from fairypptx.slides import Slides
from fairypptx.shape_range import ShapeRange
from fairypptx.table import Table
from fairypptx.core.resolvers import resolve_shapes
from fairypptx.states.shape import ShapeStateModel
from fairypptx.states.context import Context
from fairypptx.constants import msoTextOrientationHorizontal
from fairypptx.enums import MsoShapeType
from pywintypes import com_error 
from typing import cast
from PIL import Image
import numpy as np

def test_auto_shape_type():
    shape = Shape.make(1)
    text = "Hello, world!"
    shape.text = text
    shape.left = 132
    model = ShapeStateModel.from_entity(shape)
    context = Context(slide=shape.slide)
    other = model.create_entity(context)
    assert other.left == shape.left 
    assert other.text == shape.text 

def test_table_shape():
    array = np.random.normal(size=(4, 2))
    table = Table.make(array)
    shape = table.shape
    ShapeStateModel.from_entity(shape)
    context = Context(slide=shape.slide)
    model = ShapeStateModel.from_entity(shape)
    other = cast(TableShape, model.create_entity(context))
    assert np.all(other.table.to_numpy() == shape.table.to_numpy()) 

def test_group_shape_create():
    # For `GroupShape`, both tests for  `apply` and `create_entity` is necessary.
    child1 = Shape.make(1) 
    child1.text = "Child1"
    child2 = Shape.make(1) 
    child2.text = "Child2"
    group_shape = ShapeRange([child1, child2]).group()
    model = ShapeStateModel.from_entity(group_shape)

    context = Context(slide=group_shape.slide)
    other = cast(GroupShape, model.create_entity(context))
    group_shape = cast(GroupShape, group_shape)

    assert len(other.children) == len(group_shape.children)
    assert other.children[0].text == child1.text
    assert other.children[1].text == child2.text

def test_group_shape_apply():
    # For `GroupShape`, both tests for  `apply` and `create_entity` is necessary.
    child1 = Shape.make(1) 
    child1.text = "AAA1"
    child2 = Shape.make(1) 
    child2.text = "AAA2"
    group_shape = ShapeRange([child1, child2]).group()
    model = ShapeStateModel.from_entity(group_shape)
    text1 = model.impl.children[0].text_frame.api_model.text_range.runs[0].text  #type: ignore
    text2 = model.impl.children[1].text_frame.api_model.text_range.runs[0].text  #type: ignore
    assert {text1, text2} == {"AAA1", "AAA2"}
    model.apply(group_shape)


def test_picture_shape():
    # For `GroupShape`, both tests for  `apply` and `create_entity` is necessary.
    orig_array = np.random.randint(0, 256, (256, 512, 4), dtype=np.uint8)
    orig_image = Image.fromarray(orig_array, mode="RGBA")
    orig_shape = Shape.make(orig_image)
    model = ShapeStateModel.from_entity(orig_shape)

    array = np.ones(shape=(256, 512, 4), dtype=np.uint8) * 255
    image = Image.fromarray(array, mode="RGBA")
    shape = Shape.make(image)
    shape = model.apply(shape)
    assert np.allclose(np.array(orig_shape.to_image()), np.array(shape.to_image()), atol=1)

def test_textbox_shape():
    shapes_api = resolve_shapes()
    shape_api = shapes_api.AddTextbox(msoTextOrientationHorizontal, Left=0, Top=0, Width=100, Height=100) 
    shape = Shape(shape_api)
    shape.text = "Hello"
    shape.fill.color = (255, 0, 0)
    shape.line.weight = 5
    model = ShapeStateModel.from_entity(shape)

    context = Context(slide=Slide(shapes_api.Parent))
    g_shape = model.create_entity(context)
    assert shape.box == g_shape.box
    assert shape.text == g_shape.text
    assert shape.fill.color == g_shape.fill.color
    assert shape.line.weight == g_shape.line.weight

def test_placeholder_shape():
    targets = []
    for ind in range(1, 32):
        def _is_valid_target(shape) -> bool:
            if shape.api.Type == MsoShapeType.PlaceHolder:
                try:
                    shape.api.TextFrame.TextRange.Text = ""
                except com_error:
                    return False
            return True
        slide = Slides().add(layout=ind)
        targets = [shape for shape in slide.shapes if _is_valid_target(shape)]
        if targets:
            break
    if not targets:
        assert False, "Environment of powerpoint file is not appropriate."
    target = targets[0]
    assert target.api.Type == MsoShapeType.PlaceHolder
    contained_type = target.api.PlaceholderFormat.ContainedType
    target.text = "Hello, PlaceHolder."
    model = ShapeStateModel.from_entity(target)
    g_shape = model.create_entity(Context())
    assert g_shape.api.Type == contained_type, "ContainedType is unwrapped."
    assert g_shape.text == target.text, "Text is recovered."
    while 1 < len(Slides()):
        Slides().delete(len(Slides()) - 1)

def test_line_shape():
    shape = Shape(Shapes().api.AddLine(10, 20, 50, 80))
    shape.line = 5
    model = ShapeStateModel.from_entity(shape)
    context = Context(slide=shape.slide)
    other = model.create_entity(context)
    assert other.box == shape.box
    assert other.line == shape.line

def test_connector_best_effort_unresolved():
    slide = Slide()
    shapes = slide.shapes

    a = Shape(shapes.api.AddShape(1, 0, 0, 100, 100))
    b = Shape(shapes.api.AddShape(1, 200, 0, 100, 100))

    line = Shape(shapes.api.AddLine(0, 0, 200, 0))
    line.api.ConnectorFormat.BeginConnect(a.api, 1)
    line.api.ConnectorFormat.EndConnect(b.api, 1)

    model = ShapeStateModel.from_entity(line)
    context = Context(slide=Slide())  # 別スライド＝解決不能
    other = model.create_entity(context)

    assert other is not None


def test_connector_resolved():
    slide = Slide()
    shapes = slide.shapes

    a = Shape(shapes.api.AddShape(1, 0, 0, 100, 100))
    b = Shape(shapes.api.AddShape(1, 200, 0, 100, 100))

    line = Shape(shapes.api.AddLine(0, 0, 200, 0))
    line.api.ConnectorFormat.BeginConnect(a.api, 1)
    line.api.ConnectorFormat.EndConnect(b.api, 1)

    model = ShapeStateModel.from_entity(line)

    new_slide = Slide()
    context = Context(slide=new_slide)

    # 先に Shape を作って mapping を用意
    new_a = Shape(new_slide.shapes.api.AddShape(1, 0, 0, 100, 100))
    new_b = Shape(new_slide.shapes.api.AddShape(1, 200, 0, 100, 100))

    context.update_id_mapping(a.id, new_a.id)
    context.update_id_mapping(b.id, new_b.id)

    other = model.create_entity(context)

    assert other.api.Connector
    

if __name__ == "__main__":
    pass

