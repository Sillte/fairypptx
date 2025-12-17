from fairypptx.shape import Shape, TableShape, GroupShape
from fairypptx.shape_range import ShapeRange
from fairypptx.table import Table
from fairypptx.states.shape import ShapeStateModel
from fairypptx.states.context import Context
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
    child1.text = "Child1"
    child2 = Shape.make(1) 
    child2.text = "Child2"
    group_shape = ShapeRange([child1, child2]).group()
    model = ShapeStateModel.from_entity(group_shape)
    model.impl.children[0].text_frame.api_model.text_range.runs[0].text = "AAA1"  #type: ignore
    model.impl.children[1].text_frame.api_model.text_range.runs[0].text = "AAA2"  #type: ignore
    model.apply(group_shape)
    assert group_shape.children[0].text == "AAA1" #type: ignore
    assert group_shape.children[1].text == "AAA2" #type: ignore

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



if __name__ == "__main__":
    pass

