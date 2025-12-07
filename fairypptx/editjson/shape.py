from pydantic import BaseModel 
from typing import Self
from fairypptx import Shape
from fairypptx.shape import Box
from typing import Protocol
from fairypptx.editjson.protocols import EditParamProtocol
from fairypptx.editjson.line_format import NaiveLineFormatStyle
from fairypptx.editjson.fill_format import NaiveFillFormatStyle
from fairypptx.editjson.text_range import NaiveTextRangeParagraphStyle

# * Generate the parameters for `ParamItself`.
# * Apply the generate params for Shape. 

class ShapeLocationParam(BaseModel):
    box: Box  # (Note that this is jsonable).

    @classmethod
    def from_entity(cls, entity: Shape) -> Self:
        shape = entity
        return cls(box=shape.box)


    def apply(self, entity: Shape) -> Shape:
        return entity


class NaiveShapeStyle(BaseModel):
    line: NaiveLineFormatStyle
    fill: NaiveFillFormatStyle
    textrange: NaiveTextRangeParagraphStyle

    @classmethod
    def from_entity(cls, entity: Shape) -> Self:
        shape = entity
        line = NaiveLineFormatStyle.from_entity(shape.line)
        fill = NaiveFillFormatStyle.from_entity(shape.fill)
        textrange = NaiveTextRangeParagraphStyle.from_entity(shape.textrange)
        return cls(line=line, fill=fill, textrange=textrange)


    def apply(self, entity: Shape) -> Shape:
        shape = entity
        self.line.apply(shape.line)
        self.fill.apply(shape.fill)
        self.textrange.apply(shape.textrange)
        return shape


if __name__ == "__main__":
    from fairypptx import Shape  
    shape = Shape()
    target = NaiveTextRangeParagraphStyle.from_entity(shape.textrange)
    target.apply(shape.textrange)
    data = target.model_dump_json()
    print(data)
    #import time 
    #for _ in range(20):
    #    print(_)
    #    time.sleep(2)

