from fairypptx.core.models import BaseApiModel
from fairypptx.core.types import COMObject
from fairypptx.core.utils import crude_api_write, crude_api_read
from fairypptx.constants import ppBulletNone, ppBulletUnnumbered
from pydantic import BaseModel
from typing import Any, ClassVar, Literal, Mapping, Self, Sequence
from fairypptx.apis.font.api_model import FontApiModel

def _to_bool(mso_number: int) -> bool:
    return not (mso_number == 0)

class BulletFontSetting(BaseApiModel):
    font: FontApiModel | None = None
    use_text_font: bool = False
    use_text_color: bool = False

    @classmethod
    def from_api(cls, api: COMObject) -> Self:
        use_text_font = _to_bool(api.UseTextFont)
        use_text_color = _to_bool(api.UseTextColor)
        if cls._is_necessary_font(use_text_font, use_text_color):
            font = FontApiModel.from_api(api.Font)
        else:
            font = None
        return cls(font=font, use_text_font=use_text_font, use_text_color=use_text_color)

    def apply_api(self, api: COMObject):
        if self.font:
            self.font.apply_api(api.Font)
        else:
            pass
        if self.use_text_color:
            api.UseTextColor = self.use_text_color
        if self.use_text_font:
            api.UseTextFont = self.use_text_font

    @classmethod
    def _is_necessary_font(cls, use_text_font: bool, use_text_color: bool) -> bool:
        return (not use_text_font) or (not use_text_color)


class BulletFormatApiModel(BaseApiModel):
    type: int
    api_data: Mapping[str, Any]
    font_setting: BulletFontSetting
    character: int | None = None

    _keys: ClassVar[Sequence[str]] = [
        "Visible",
        "RelativeSize",
        ]

    @classmethod
    def from_api(cls, api: COMObject) -> Self:
        type = api.Type
        data = crude_api_read(api, cls._keys)
        font_setting = BulletFontSetting.from_api(api)
        if api.Type == ppBulletUnnumbered:
            character = api.Character
        else:
            character = None
        return cls(api_data=data, font_setting=font_setting, type=type, character=character)

    def apply_api(self, api: COMObject):
        api.Type = self.type
        if self.type != ppBulletNone:
            if self.character is not None:
                api.Character = self.character
            self.font_setting.apply_api(api)
            crude_api_write(api, self.api_data)


