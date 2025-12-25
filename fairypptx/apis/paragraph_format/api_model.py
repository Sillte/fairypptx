from fairypptx import constants
from fairypptx.core.models import BaseApiModel
from fairypptx.core.types import COMObject
from fairypptx.core.utils import crude_api_read, crude_api_write, remove_invalidity
from fairypptx.object_utils import to_api2
from fairypptx.apis.bullet_format.api_model import BulletFormatApiModel

from collections.abc import Mapping, Sequence
from typing import Any, ClassVar, Mapping, Self, Sequence


class ParagraphFormatApiModel(BaseApiModel):
    """
    Note
    -------------------------------------
    Curently, About `data`, the order of key is important
    since some keys (I infer ``Bullet`.Character`?) change the other properties implicitly.
    This knowledge must be also taken care by users to customize.
    [TODO] You can modify this. See ``FillFormat``.

    Wondering...
    -----------------------------------------
    BulletFormat is introduced or not...
    * https://docs.microsoft.com/ja-jp/office/vba/api/powerpoint.bulletformat.number
    """

    api_data: Mapping[str, Any]
    api2_data: Mapping[str, Any] = {}
    bullet: BulletFormatApiModel | None = None

    _common_keys: ClassVar[Sequence[str]] = [
            "FarEastLineBreakControl",
            "Alignment",
            "BaseLineAlignment",
            "HangingPunctuation",
            "SpaceAfter",
            "SpaceBefore",
            "SpaceWithin"]


    _api2_keys: ClassVar[Sequence[str]] = [
        "FirstLineIndent",
        "LeftIndent",
        ]


    @classmethod
    def from_api(cls, api: COMObject) -> Self:
        """Generate itself from the entity of `fairpptx.PPTXObject`
        """

        api2 = to_api2(api)
        keys = set(cls._common_keys)

        if api.Bullet.Type != constants.ppBulletNone:
            bullet = BulletFormatApiModel.from_api(api.Bullet)
        else:
            bullet = None

        #if api.Bullet.Type != constants.ppBulletUnnumbered:
        #    keys -= {"Bullet.Character", "Bullet.Font.Name"}

        api_data = crude_api_read(api, list(keys))
        api2_data = crude_api_read(api2, cls._api2_keys)

        api_data = remove_invalidity(api, api_data)
        api2_data = remove_invalidity(api2, api2_data)

        return cls(api_data=api_data, api2_data=api2_data, bullet=bullet)

    def apply_api(self, api: COMObject) -> COMObject:
        crude_api_write(api, self.api_data)
        api2 = to_api2(api)
        crude_api_write(api2, self.api2_data)

        if self.bullet:
            self.bullet.apply_api(api.Bullet)
        else:
            api.Bullet.Type = constants.ppBulletNone

        return api

