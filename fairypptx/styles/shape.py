import warnings 
from pydantic import BaseModel 
from typing import Self
from fairypptx import Shape
from fairypptx.shape import Box
from typing import Literal
from fairypptx.styles.protocols import StyleModelProtocol
from fairypptx.styles.line_format import NaiveLineFormatStyle
from fairypptx.styles.fill_format import NaiveFillFormatStyle
from fairypptx.styles.text_range import NaiveTextRangeParagraphStyle
from fairypptx.styles.text_frame import NaiveTextFrameStyle

from fairypptx.core.utils import get_discriminator_mapping
from fairypptx.enums import MsoShapeType


class ShapeLocationParam(BaseModel):
    box: Box  # (Note that this is jsonable).

    @classmethod
    def from_entity(cls, entity: Shape) -> Self:
        shape = entity
        return cls(box=shape.box)


    def apply(self, entity: Shape) -> Shape:
        return entity

class AutoShapeStyle(BaseModel):
    type: Literal[MsoShapeType.AutoShape] = MsoShapeType.AutoShape
    auto_shape_type: int 
    line: NaiveLineFormatStyle
    fill: NaiveFillFormatStyle
    text_frame: NaiveTextFrameStyle
    is_tight: bool

    @classmethod
    def from_entity(cls, entity: Shape) -> Self:
        shape = entity
        line = NaiveLineFormatStyle.from_entity(shape.line)
        fill = NaiveFillFormatStyle.from_entity(shape.fill)
        text_frame = NaiveTextFrameStyle.from_entity(shape.text_frame)
        is_tight = entity.is_tight()
        return cls(line=line, fill=fill, text_frame=text_frame, auto_shape_type=entity.api.AutoShapeType, is_tight=is_tight)

    def apply(self, entity: Shape) -> Shape:
        shape = entity
        shape.api.AutoShapeType = self.auto_shape_type
        self.line.apply(shape.line)
        self.fill.apply(shape.fill)
        self.text_frame.apply(shape.text_frame)
        if self.is_tight:
            entity.tighten()
        return shape


class TextBoxStyle(BaseModel):
    type: Literal[MsoShapeType.TextBox] = MsoShapeType.TextBox
    text_frame: NaiveTextFrameStyle

    @classmethod
    def from_entity(cls, entity: Shape) -> Self:
        shape = entity
        text_frame = NaiveTextFrameStyle.from_entity(shape.text_frame)
        return cls(text_frame=text_frame)

    def apply(self, entity: Shape) -> Shape:
        shape = entity
        self.text_frame.apply(shape.text_frame)
        return shape
    

class FallbackShapeStyle(BaseModel):
    type: int

    @classmethod
    def from_entity(cls, entity: Shape) -> Self:
        return cls(type=entity.api.Type)
    def apply(self, entity: Shape) -> Shape:
        msg = f"`{self.type}` cannot be handled."
        warnings.warn(msg)
        return entity

ShapeStyle = AutoShapeStyle | TextBoxStyle


class NaiveShapeStyle(BaseModel):
    selector: ShapeStyle | FallbackShapeStyle

    @classmethod
    def from_entity(cls, entity: Shape) -> Self:
        cls_mapping = get_discriminator_mapping(ShapeStyle, "type")
        klass = cls_mapping.get(entity.api.Type)
        if klass:
            selector=klass.from_entity(entity)
        else:
            selector=FallbackShapeStyle(type=entity.api.Type)
        return cls(selector=selector)
    

    def apply(self, entity: Shape) -> Shape:
        if self.selector.type == entity.api.Type:
            self.selector.apply(entity)
        else:
            msg = f"This class is applicable when `type={self.selector.type}`, but the given is `{entity.api.Type}`" 
            raise TypeError(msg)
        return entity


if __name__ == "__main__":
    from fairypptx import Shape  
    shape = Shape()
    target = NaiveTextRangeParagraphStyle.from_entity(shape.textrange)
    target.apply(shape.textrange)

