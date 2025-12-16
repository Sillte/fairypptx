from fairypptx.core.models import BaseApiModel
from fairypptx.core.types import COMObject
from pywintypes import com_error

from fairypptx.core.utils import crude_api_read, crude_api_write

from typing import Self, Mapping, Any
from fairypptx.apis.text_range import TextRangeApiModel
from fairypptx.object_utils import to_api2
from fairypptx.object_utils import setattr as f_setattr


common_keys: list[str] = [
    # 1. 流し込み/向きを決定するプロパティを最初に
    "Orientation", 

    # 2. 余白
    "MarginLeft",
    "MarginRight", 
    "MarginTop",
    "MarginBottom",
    
    # 3. 配置を最後に（Orientationの設定が確定してから）
    "VerticalAnchor",
    "HorizontalAnchor",
]

special_keys: list[str] = ["AutoSize"]


def to_style_api_data(api: COMObject) -> Mapping[str, Any]:
    data =  crude_api_read(api, common_keys)
    data.update(crude_api_read(api, special_keys))
    return data

def apply_style_api_data(api: COMObject, data: Mapping[str, Any]):
    special_data = {key: value for key, value in data.items() if key in special_keys}
    normal_data = {key: value for key, value in data.items() if key not in special_keys}
    for key, value in special_data.items():
        try:
            f_setattr(api, key, value)
        except ValueError:
            pass
    return crude_api_write(api, normal_data)

common_keys2 = ["WordArtformat"]

def to_style_api2_data(api: COMObject) -> Mapping[str, Any]:
    api2 = to_api2(api)
    return crude_api_read(api2, common_keys)

def apply_style_api2_data(api: COMObject, data: Mapping[str, Any]):
    api2 = to_api2(api)
    return crude_api_write(api2, data)


class TextFrameApiModel(BaseApiModel):
    api_data: Mapping[str, Any]
    api2_data: Mapping[str, Any]
    text_range: TextRangeApiModel 


    @classmethod
    def from_api(cls, api: COMObject) -> Self:
        api2_data = to_style_api2_data(api)
        api_data = to_style_api_data(api)
        tr = TextRangeApiModel.from_api(api.TextRange)
        return cls(text_range=tr, api_data=api_data, api2_data=api2_data)

    def apply_api(self, api: COMObject) -> None:
        # The order of `api2` and `api` is important.
        apply_style_api2_data(api, self.api2_data)
        apply_style_api_data(api, self.api_data)
        self.text_range.apply_api(api.TextRange)
