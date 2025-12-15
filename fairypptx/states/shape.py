from fairypptx.core.utils import get_discriminator_mapping
from fairypptx.states.models import BaseStateModel
from fairypptx.shape import Shape
from fairypptx.box import Box 
from pydantic import Field
from typing import Annotated, Self, Literal
from fairypptx.styles.line_format import NaiveLineFormatStyle
from fairypptx.styles.fill_format import NaiveFillFormatStyle
from fairypptx.enums import MsoShapeType

class AutoShapeStateModel(BaseStateModel):
    type: Annotated[Literal[MsoShapeType.AutoShape], Field(description="Type of Shape")] = MsoShapeType.AutoShape
    box: Annotated[Box, Field(description="Represents the position of the shape")]  # (Note that this is jsonable).
    auto_shape_type: Annotated[int, Field(description="Represents MSOAutoShapeType.")]
    line: Annotated[NaiveLineFormatStyle, Field(description="Represents the format of `Line` around the Shape.")]
    fill: Annotated[NaiveFillFormatStyle, Field(description="Represents the format of `Fill` of the Shape.")]

    @classmethod
    def from_entity(cls, entity: Shape) -> Self:
        shape = entity
        return cls(box=shape.box,
                   id=shape.id, 
                   auto_shape_type=shape.api.AutoShapeType, 
                   line=NaiveLineFormatStyle.from_entity(shape.line), 
                   fill=NaiveFillFormatStyle.from_entity(shape.fill)
                   )

    def apply(self, entity: Shape):
        shape = entity
        shape.box = self.box
        shape.api.AutoShapeType = self.auto_shape_type
        self.line.apply(shape.line)
        self.fill.apply(shape.fill)


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

    def apply(self, entity: Shape):
        shape = entity
        shape.box = self.box

type ShapeStateModelImpl = AutoShapeStateModel 

class ShapeStateModel(BaseStateModel):
    impl: Annotated[ShapeStateModelImpl | FallbackShapeStateModel, Field(description="Based on `Type`, appropriate class is selected.")]

    @classmethod
    def from_entity(cls, entity: Shape) -> Self:
        cls_mapping = get_discriminator_mapping(ShapeStateModel, "type")
        klass = cls_mapping[entity.api.Type]
        impl = klass.from_entity(entity)
        return cls(impl=impl, id=impl.id)

    def apply(self, entity: Shape):
        self.impl.apply(entity)

