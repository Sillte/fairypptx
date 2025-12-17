"""Text Frame API layer exports."""
from fairypptx.apis.text_range.api_model import TextRangeApiModel, normalize_paragraph_breaks
from fairypptx.apis.text_range.applicator import TextRangeApplicator

__all__ = ["TextRangeApiModel", "TextRangeApplicator", "normalize_paragraph_breaks"]

