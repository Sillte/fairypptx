

"""Paragraph Format API Applicator: Apply ParagraphFormatApiModel to COMObject.

This module applies ParagraphFormatApiModel instances (Pydantic models) to Win32 COM 
Paragraph objects.

No custom conversion strategy is needed because:
	- Paragraph API has a single schema (no type variants)
	- Value types (Alignment, Indent, etc.) are handled by to_api2() in api_model
	- Direct COMObject → COMObject mapping via ApiApplicator base class suffices

Design note:
	- apply_custom is None, so ApiApplicator.apply() uses only Pydantic model conversion
	- Key order preservation in 'data' mapping is critical for PowerPoint API compatibility
"""

from fairypptx.core.models import ApiApplicator
from fairypptx.apis.paragraph_format.api_model import ParagraphFormatApiModel


ParagraphFormatApplicator = ApiApplicator(ParagraphFormatApiModel, None)
