from fairypptx.apis.paragraph_format.api_model import ParagraphFormatApiModel
from fairypptx.apis.paragraph_format.applicator import ParagraphFormatApplicator
from fairypptx.core.types import COMObject
from fairypptx.object_utils import to_api2

from typing import Literal, Self, Annotated, cast

from fairypptx.object_utils import to_api2, is_object


class ParagraphFormat:
    """Represents the Font Information. """

    def __init__(self, api: COMObject) -> None:
        if isinstance(api, ParagraphFormat):
            api = api.api
        assert is_object(api)
        self._api = api
        
    @property
    def api(self) -> COMObject:
        return self._api

    @property
    def api2(self) -> COMObject:
        return to_api2(self._api)
    

class ParagraphFormatProperty:
    def __get__(self, parent: COMObject, objtype=None):
        return ParagraphFormat(parent.api.ParagraphFormat)


    def __set__(self, shape, value):
        ParagraphFormatApplicator.apply(shape.api.ParagraphFormat, value)



if __name__ == "__main__":
    pass
