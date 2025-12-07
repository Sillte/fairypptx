from pydantic import BaseModel, JsonValue, TypeAdapter, Field

from enum import IntEnum
from typing import Any, Mapping, Literal, ClassVar, Sequence, Self, Annotated, cast
from fairypptx._shape import FillFormat
from fairypptx._shape import LineFormat
from fairypptx import constants

from fairypptx.constants import msoFillSolid, msoFillPatterned, msoFillGradient
from fairypptx.object_utils import setattr, getattr
from fairypptx.editjson.utils import get_discriminator_mapping, CrudeApiAccesssor, crude_api_read, f_setattr, crude_api_write
from pprint import pprint
from fairypptx.enums import MsoFillType
from fairypptx.editjson.protocols import ApiApplyBaseModel, EditParamProtocol
from pywintypes import com_error


class NaiveLineFormatStyle(BaseModel):
    body: Mapping[str, Any]

    _common_keys: ClassVar[Sequence[str]] = [
            "BackColor.RGB",
            "DashStyle",
            "ForeColor.RGB",
            "InsetPen",
            "Pattern",
            "Transparency",
            "Visible",
            "Weight",
            "Style"]

    @classmethod
    def from_entity(cls, entity: LineFormat) -> Self:
        """Generate itself from the entity of `fairpptx.PPTXObject`
        """
        data = dict()
        data["Style"] = constants.msoLineSingle
        data["ForeColor.RGB"] = 0
        data["Visible"] = constants.msoTrue
        data["Transparency"] = 0

        api = entity.api
        keys = list(cls._common_keys)

        if getattr(api, "BeginArrowheadStyle") != constants.msoArrowheadNone:
            keys += ["BeginArrowheadStyle", "BeginArrowheadLength", "BeginArrowheadWidth"]
        if getattr(api, "EndArrowheadStyle") != constants.msoArrowheadNone:
            keys += ["EndArrowheadStyle", "EndArrowheadLength", "EndArrowheadWidth"]

        # (2021/05/19)
        # For some keys, invalid values are initially set.
        # `setattr` of invalid values raises `ValueError`.
        # Using this knowledge, remove the not-apt keys.

        data.update(crude_api_read(api, keys))
        remove_keys = set()
        for key, value in data.items():
            try:
                f_setattr(api, key, value)
            except ValueError:
                remove_keys.add(key)
        data = {key: value for key, value in data.items() if key not in remove_keys}

        return cls(body=data)


    def apply(self, entity: LineFormat) -> LineFormat:
        """Apply this edit param to the entity.
        """
        crude_api_write(entity.api, self.body)
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


