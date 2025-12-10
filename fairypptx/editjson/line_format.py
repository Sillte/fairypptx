from pydantic import BaseModel, JsonValue, TypeAdapter, Field

from enum import IntEnum
from typing import Any, Mapping, Literal, ClassVar, Sequence, Self, Annotated, cast
from fairypptx._shape import FillFormat
from fairypptx._shape import LineFormat
from fairypptx import constants

from fairypptx.constants import msoFillSolid, msoFillPatterned, msoFillGradient
from fairypptx.core.models import ApiBridgeBaseModel
from fairypptx.core.utils import CrudeApiAccesssor, crude_api_read, crude_api_write, get_discriminator_mapping, remove_invalidity
from pprint import pprint
from fairypptx.enums import MsoFillType
from fairypptx.editjson.protocols import EditParamProtocol
from pywintypes import com_error
from fairypptx.line_format import LineFormatApiBridge


class NaiveLineFormatStyle(BaseModel):
    api_bridge: LineFormatApiBridge

    @classmethod
    def from_entity(cls, entity: LineFormat) -> Self:
        api_bridge = LineFormatApiBridge.from_api(entity.api)
        return cls(api_bridge=api_bridge)

    def apply(self, entity: LineFormat) -> LineFormat:
        """Apply this edit param to the entity.
        """
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


