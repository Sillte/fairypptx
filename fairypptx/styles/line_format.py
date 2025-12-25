from pydantic import BaseModel

from typing import Any, Mapping, Literal, ClassVar, Sequence, Self, Annotated, cast
from fairypptx.line_format import LineFormat

from fairypptx.apis.line_format.api_model import LineFormatApiModel


class NaiveLineFormatStyle(BaseModel):
    api_bridge: LineFormatApiModel

    @classmethod
    def from_entity(cls, entity: LineFormat) -> Self:
        api_bridge = LineFormatApiModel.from_api(entity.api)
        return cls(api_bridge=api_bridge)

    def apply(self, entity: LineFormat) -> LineFormat:
        """Apply this edit param to the entity.
        """
        print("model", self.api_bridge)
        self.api_bridge.apply_api(entity.api)
        return entity


if __name__ == "__main__":
    from fairypptx import Shape  
    shape = Shape()
    target = NaiveLineFormatStyle.from_entity(shape.line)
    data = target.model_dump_json()
    import time 
    for _ in range(20):
        print(_)
        time.sleep(2)

    target.apply(shape.line)


