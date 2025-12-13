from fairypptx.core.resolvers import resolve_textframe
from fairypptx.core.types import COMObject, PPTXObjectProtocol  

from fairypptx.apis.text_frame import TextFrameApplicator 

from fairypptx.text_range import TextRange, TextRangeProperty

class TextFrame:
    text_range = TextRangeProperty()
    textrange = TextRangeProperty() # backward-compatibility.
 
    def __init__(self, arg):
        self._api = resolve_textframe(arg)

    @property
    def api(self) -> COMObject:
        return self._api


    @property
    def api2(self) -> COMObject:
        return self._api.Parent.TextFrame2


class TextFrameProperty:
    def __get__(self, parent: PPTXObjectProtocol, objtype=None) -> TextFrame:
        return TextFrame(parent.api.TextFrame)

    def __set__(self, parent: PPTXObjectProtocol, value: str | TextRange) -> None:
        TextFrameApplicator.apply(parent.api.TextFrame, value)
