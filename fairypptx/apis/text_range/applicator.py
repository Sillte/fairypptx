from fairypptx import constants
from fairypptx import object_utils
from fairypptx.apis.line_format.api_model import LineFormatApiModel
from fairypptx.color import Color, ColorLike
from fairypptx.core.types import COMObject
from fairypptx.core.models import ApiApplicator
from fairypptx.apis.text_range.api_model import TextRangeApiModel

def apply_custom(api: COMObject, value: str | COMObject):
    if isinstance(value, str): 
        api.Text = value
    else:
        raise ValueError("Cannot handle `value`.")


TextRangeApplicator = ApiApplicator(TextRangeApiModel, apply_custom)
