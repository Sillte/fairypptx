from pydantic import BaseModel, JsonValue, TypeAdapter, Field

from enum import IntEnum
from typing import Any, Mapping, Literal, ClassVar, Sequence, Self, Annotated, cast
from fairypptx._shape import FillFormat
from fairypptx._shape import LineFormat
from fairypptx import constants

from fairypptx.constants import msoFillSolid, msoFillPatterned, msoFillGradient
from fairypptx.core.models import BaseApiModel
from fairypptx.core.utils import CrudeApiAccesssor, crude_api_read, crude_api_write, get_discriminator_mapping, remove_invalidity
from fairypptx.object_utils import setattr, getattr
from fairypptx.paragraph_format import ParagraphFormat
from pprint import pprint
from fairypptx.enums import MsoFillType
from fairypptx.editjson.protocols import EditParamProtocol
from pywintypes import com_error



class NaiveParagraphFormatStyle(BaseModel):
    api_data: Mapping[str, Any]
    api2_data: Mapping[str, Any] = {}
    

    _common_keys: ClassVar[Sequence[str]] = [
            "FarEastLineBreakControl", "Alignment",
            "BaseLineAlignment",
            "HangingPunctuation",
            "LineRuleAfter",
            "LineRuleBefore",
            "LineRuleWithin",
            "SpaceAfter",
            "SpaceBefore",
            "SpaceWithin"]

    # The order is very important!
    # Especially, `Type` and `Visible`!.
    _bullet_keys: ClassVar[Sequence[str]] = [
        "Bullet.Type",
        "Bullet.Visible",
        "Bullet.Character",
        "Bullet.Font.Name",
        ]

    _api2_keys: ClassVar[Sequence[str]] = [
        "FirstLineIndent",
        "LeftIndent",
        ]


    @classmethod
    def from_entity(cls, entity: ParagraphFormat) -> Self:
        """Generate itself from the entity of `fairpptx.PPTXObject`
        """
        api = entity.api
        assert api 
        api2 = entity.api2
        keys = set(cls._common_keys) | set(cls._bullet_keys)

        if api.Bullet.Type != constants.ppBulletUnnumbered:
            keys -= {"Bullet.Character", "Bullet.Font.Name"}

        api_data = crude_api_read(api, list(keys))
        api2_data = crude_api_read(api2, cls._api2_keys)

        api_data = remove_invalidity(api, api_data)
        api2_data = remove_invalidity(api2, api2_data)

        return cls(api_data=api_data, api2_data=api2_data)

    def apply(self, entity: ParagraphFormat) -> ParagraphFormat:
        """Apply this edit param to the entity.

        Sets both `api` (v1) and `api2` (v2) properties.
        Skips keys with None values to preserve defaults.
        """
        api = entity.api
        api2 = entity.api2

        if not api or not api2:
            raise ValueError("ParagraphFormat entity must have both api and api2 properties.")

        crude_api_write(api, self.api_data)
        crude_api_write(api2, self.api2_data)
        return entity


if __name__ == "__main__":
    from fairypptx import Shape  
    shape = Shape()
    target = NaiveParagraphFormatStyle.from_entity(shape.textrange.paragraph_format)
    data = target.model_dump_json()
    import time 
    for _ in range(20):
        print(_)
        time.sleep(2)

    target.apply(shape.textrange.paragraph_format)


