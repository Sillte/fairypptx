
from fairypptx.states.models import BaseValueModel
from fairypptx.table import Table 
from pydantic import Field
from typing import Annotated, Self
from fairypptx.apis.text_frame import TextFrameApiModel
from fairypptx.text_frame import TextFrame


class TextFrameValueModel(BaseValueModel):
    api_model: Annotated[TextFrameApiModel, Field(description="TextFrame. It contains the text of Shape.")]


    @classmethod
    def from_object(cls, object: TextFrame) -> Self:
        text_frame = object
        return cls(api_model=TextFrameApiModel.from_api(text_frame.api))


    def apply(self, object: TextFrame):
        text_frame = object
        self.api_model.apply_api(text_frame.api)


