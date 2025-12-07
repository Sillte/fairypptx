from pydantic import BaseModel 
from typing import Self
from fairypptx import Shape
from fairypptx.shape import Box
from typing import Protocol
from fairypptx.editjson.protocols import EditParamProtocol

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


if __name__ == "__main__":

    box = Box(1,2,3,4)

    param = ShapeLocationParam(box=box)
    print(param.model_dump())





if __name__ == "__main__":
    pass
