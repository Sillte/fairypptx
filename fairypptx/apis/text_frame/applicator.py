from fairypptx import constants
from fairypptx.apis.line_format.api_model import LineFormatApiModel
from fairypptx.color import Color, ColorLike
from fairypptx.core.types import COMObject
from fairypptx.core.models import ApiApplicator
from fairypptx.apis.text_frame.api_model import TextFrameApiModel


TextFrameApplicator = ApiApplicator(TextFrameApiModel, None)
