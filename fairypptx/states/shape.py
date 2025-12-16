from fairypptx.core.utils import get_discriminator_mapping
from fairypptx.states.models import BaseStateModel
from fairypptx.states.context import Context
from fairypptx.shape import Shape, TableShape
from fairypptx.box import Box 
from pydantic import Field
from typing import Annotated, Self, Literal, cast
from fairypptx.styles.line_format import NaiveLineFormatStyle
from fairypptx.styles.fill_format import NaiveFillFormatStyle
from fairypptx.states.table import TableValueModel
from fairypptx.states.text_frame import TextFrameValueModel
from fairypptx.table import Table
from fairypptx.enums import MsoShapeType

class AutoShapeStateModel(BaseStateModel):
    type: Annotated[Literal[MsoShapeType.AutoShape], Field(description="Type of Shape")] = MsoShapeType.AutoShape
    box: Annotated[Box, Field(description="Represents the position of the shape")]  # (Note that this is jsonable).
    auto_shape_type: Annotated[int, Field(description="Represents MSOAutoShapeType.")]
    line: Annotated[NaiveLineFormatStyle, Field(description="Represents the format of `Line` around the Shape.")]
    fill: Annotated[NaiveFillFormatStyle, Field(description="Represents the format of `Fill` of the Shape.")]
    text_frame: Annotated[TextFrameValueModel, Field(description="Represents the texts of the Shape.")]

    @classmethod
    def from_entity(cls, entity: Shape) -> Self:
        shape = entity
        return cls(box=shape.box,
                   id=shape.id, 
                   auto_shape_type=shape.api.AutoShapeType, 
                   line=NaiveLineFormatStyle.from_entity(shape.line), 
                   fill=NaiveFillFormatStyle.from_entity(shape.fill),
                   text_frame=TextFrameValueModel.from_object(shape.text_frame)
                   )

    def create_entity(self, context: Context) -> Shape: 
        shapes = context.shapes
        shape = shapes.add(shape_type=self.auto_shape_type)
        self.apply(shape)
        return shape

    def apply(self, entity: Shape):
        shape = entity
        shape.box = self.box
        shape.api.AutoShapeType = self.auto_shape_type
        self.line.apply(shape.line)
        self.fill.apply(shape.fill)
        self.text_frame.apply(shape.text_frame)

class TableShapeStateModel(BaseStateModel):
    type: Annotated[Literal[MsoShapeType.Table], Field(description="Type of Shape")] = MsoShapeType.Table
    box: Annotated[Box, Field(description="Represents the position of the shape")]  # (Note that this is jsonable).
    table: Annotated[TableValueModel, Field(description="Table of the Shape")]
    @classmethod
    def from_entity(cls, entity: Shape) -> Self:
        shape = cast(TableShape, entity)
        return cls(box=shape.box,
                   id=shape.id, 
                   table=TableValueModel.from_object(shape.table)
                   )

    def create_entity(self, context: Context) -> Shape:
        shapes = context.shapes
        n_rows = self.table.n_rows
        n_columns = self.table.n_columns
        shape_api = shapes.api.AddTable(NumRows=n_rows, NumColumns=n_columns)
        table = Table(shape_api.Table)
        shape = table.shape
        self.apply(shape)
        return table.shape
 
    def apply(self, entity: Shape):
        shape = cast(TableShape, entity)
        shape.box = self.box
        self.table.apply(shape.table)



class FallbackShapeStateModel(BaseStateModel):
    box: Annotated[Box, Field(description="Represents the position of the shape")] 
    type: int 

    @classmethod
    def from_entity(cls, entity: Shape) -> Self:
        shape = entity
        return cls(box=shape.box,
                   id=shape.id,
                   type=shape.api.Type
                   )

    def create_entity(self, context: Context) -> Shape:
        shapes = context.shapes
        shape = shapes.add(1)
        shape.text = f"Created, but `{self.type}` cannnot be handled."
        self.apply(shape)
        return shape

    def apply(self, entity: Shape):
        shape = entity
        shape.box = self.box

ShapeStateModelImpl = AutoShapeStateModel | TableShapeStateModel

class ShapeStateModel(BaseStateModel):
    impl: Annotated[ShapeStateModelImpl | FallbackShapeStateModel, Field(description="Based on `Type`, appropriate class is selected.")]

    def create_entity(self, context: Context) -> Shape:
        return self.impl.create_entity(context)


    @classmethod
    def from_entity(cls, entity: Shape) -> Self:
        cls_mapping = get_discriminator_mapping(ShapeStateModelImpl, "type")
        print(cls_mapping, entity.api.Type)
        klass = cls_mapping[entity.api.Type]
        impl = klass.from_entity(entity)
        return cls(impl=impl, id=impl.id)

    def apply(self, entity: Shape):
        self.impl.apply(entity)


if __name__ == "__main__":
    from fairypptx import Shape
    shape_state = ShapeStateModel.from_entity(Shape())
    print(shape_state)

