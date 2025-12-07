from pydantic import BaseModel, JsonValue, TypeAdapter, Field

from typing import Self
from fairypptx._shape import FillFormat
from pprint import pprint
from fairypptx.enums import MsoFillType
from pywintypes import com_error
from fairypptx._shape.fill_format import  FillFormatApiBridge


class NaiveFillFormatStyle(BaseModel):
    api_bridge: FillFormatApiBridge

    @classmethod
    def from_entity(cls, entity: FillFormat) -> Self:
        """Generate itself from the entity of `fairpptx.PPTXObject`
        """
        api_bridge = FillFormatApiBridge.from_api(entity.api)
        return cls(api_bridge=api_bridge)

    def apply(self, entity: FillFormat) -> FillFormat:
        """Apply this edit param to 
        """
        self.api_bridge.apply_api(entity.api)
        return entity

if __name__ == "__main__":
    from fairypptx import Shape  
    shape = Shape()
    target = NaiveFillFormatStyle.from_entity(shape.fill)
    print(target.model_dump_json())