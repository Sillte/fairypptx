from fairypptx.shape import Shape, TableShape
from fairypptx.table import Table
from fairypptx.states.shape import ShapeStateModel
from fairypptx.states.context import Context
from typing import cast
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


if __name__ == "__main__":
    pass

