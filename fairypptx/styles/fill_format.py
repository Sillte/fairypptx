from pydantic import BaseModel

from typing import Self
from fairypptx.fill_format import FillFormat
from fairypptx.apis.fill_format.api_model import FillFormatApiModel
from fairypptx.enums import MsoFillType



class NaiveFillFormatStyle(BaseModel):
    api_bridge: FillFormatApiModel

    @classmethod
    def from_entity(cls, entity: FillFormat) -> Self:
        """Generate itself from the entity of `fairpptx.PPTXObject`
        """
        api_bridge = FillFormatApiModel.from_api(entity.api)
        return cls(api_bridge=api_bridge)

    @property
    def valid(self) -> bool:
        return self.api_bridge.api_data.type != MsoFillType.FillMixed

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
