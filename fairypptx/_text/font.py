from fairypptx import constants 


from pywintypes import com_error
from collections.abc import Mapping 

from fairypptx import constants
from fairypptx.object_utils import to_api2


from collections.abc import Sequence 
from fairypptx import constants
from fairypptx.color import Color

from typing import Any, Mapping, Literal, ClassVar, Sequence, Self, Annotated, cast
from fairypptx.core.models import ApiBridgeBaseModel
from fairypptx.core.utils import crude_api_read, crude_api_write, remove_invalidity
from fairypptx.core.types import COMObject

from fairypptx.object_utils import to_api2, getattr, is_object, setattr
from fairypptx.object_utils import getattr as f_getattr


class FontApiBridgeBaseModel(ApiBridgeBaseModel):
    """Naive font edit parameter that owns dict<->Font conversion.

    Responsibilities:
    - `from_entity`: read a `Font` wrapper or COM `Font` and build a plain mapping.
    - `apply`: write the mapping back to a `Font` wrapper or COM `Font`.

    Implementation notes:
    - `crude_api_read` / `crude_api_write` are used for common keys.
    - For boolean-like properties we only keep keys whose values are one of
      the MSO tri-state constants (msoCTrue/msoTrue/msoFalse) to avoid
      storing other sentinel values.
    """

    api_data: Mapping[str, Any]
    _common_keys: ClassVar[Sequence[str]] = ["Size", "Name", "Color.RGB"]
    _only_determined_keys: ClassVar[Sequence[str]] = [
        "Bold",
        "Italic",
        "Shadow",
        "Superscript",
        "Subscript",
        "Underline",
    ]

    @classmethod
    def from_api(cls, api: COMObject) -> Self:
        data: dict[str, Any] = {}
        # read common keys in bulk
        data.update(crude_api_read(api, cls._common_keys))

        # read boolean-like keys only when they look like MSO tri-state values
        for key in cls._only_determined_keys:
            try:
                value = f_getattr(api, key)
            except (com_error, AttributeError):
                continue
            if value in {constants.msoCTrue, constants.msoTrue, constants.msoFalse}:
                data[key] = value
        return cls(api_data=data)
    
    def apply_api(self, api:COMObject) -> Self:
        crude_api_write(api, self.api_data)
        return api


class Font:
    """Represents the Font Information. 
    """
    def __init__(self, api):
        if isinstance(api, Font):
            api = api.api
        assert is_object(api)
        self._api = api

    @property
    def api(self):
        return self._api

    def apply(self, other: Self) -> None:
        api_bridge = FontApiBridgeBaseModel.from_api(other.api)
        api_bridge.apply_api(self.api)

    @property
    def bold(self) -> bool:
        return self.api.Bold != constants.msoFalse

    @bold.setter
    def bold(self, value: bool | int):
        if value is True:
            self.api.Bold = constants.msoTrue
        elif value is False:
            self.api.Bold = constants.msoFalse
        else:
            self.api.Bold = value

    @property
    def underline(self):
        return self.api.Underline != constants.msoFalse

    @underline.setter
    def underline(self, value):
        if value is True:
            self.api.Underline = constants.msoTrue
        elif value is False:
            self.api.Underline = constants.msoFalse
        else:
            self.api.Underline = value

    @property
    def color(self):
        return Color(self.api.Color.RGB)

    @color.setter
    def color(self, value):
        value = Color(value)
        self.api.Color.RGB = value.as_int()
        if value.alpha < 1:
            api2 = to_api2(self.api)
            api2.Fill.Transparency = 1 - value.alpha 

    @property
    def size(self):
        return self.api.Size

    @size.setter
    def size(self, value):
        self.api.Size = value
