from pydantic import BaseModel
from fairypptx.text_frame import TextFrame 
from fairypptx.apis.text_frame.api_model import to_style_api_data, apply_style_api_data
from fairypptx.apis.text_frame.api_model import to_style_api2_data, apply_style_api2_data
from fairypptx.styles.text_range import NaiveTextRangeParagraphStyle
from typing import Self, Any, Mapping


class NaiveTextFrameStyle(BaseModel):
    api_data: Mapping[str, Any]
    api2_data: Mapping[str, Any]
    text_range_style: NaiveTextRangeParagraphStyle

    @classmethod
    def from_entity(cls, entity: TextFrame) -> Self:
        api_data = to_style_api_data(entity.api)
        api2_data = to_style_api2_data(entity.api)
        text_range_style = NaiveTextRangeParagraphStyle.from_entity(entity.text_range)
        return cls(api_data=api_data, api2_data=api2_data, text_range_style=text_range_style)
    

    def apply(self, entity: TextFrame):
        self.text_range_style.apply(entity.text_range)
        apply_style_api2_data(entity.api, self.api2_data)
        apply_style_api_data(entity.api, self.api_data)
        return entity
