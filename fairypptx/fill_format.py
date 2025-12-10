from pywintypes import com_error

from fairypptx import constants
from fairypptx.color import Color, ColorLike
from fairypptx.core.types import COMObject

from typing import Any, Mapping, Literal, ClassVar, Sequence, Self, Annotated, cast, TYPE_CHECKING
from fairypptx.core.utils import get_discriminator_mapping
from fairypptx.core.utils import CrudeApiAccesssor
from fairypptx.enums import MsoFillType
from fairypptx.core.models import ApiBridgeBaseModel
from pywintypes import com_error

if TYPE_CHECKING:
    from fairypptx.shape import Shape


class NaiveSolidFillFormat(ApiBridgeBaseModel):
    type: Literal[MsoFillType.FillSolid] = MsoFillType.FillSolid
    data: Mapping[str, Any]
    _keys: ClassVar[Sequence[str]] = ["ForeColor.RGB", "Visible", "Transparency"]
    _accessor: ClassVar[CrudeApiAccesssor] = CrudeApiAccesssor(_keys)

    def apply_api(self, api):
        self._accessor.write(api, self.data)

    @classmethod
    def from_api(cls, api) -> Self:
        api.Solid()
        data = cls._accessor.read(api)
        return cls(data=data)


class NaivePatternedFillFormat(ApiBridgeBaseModel):
    type: Literal[MsoFillType.FillPatterned] = MsoFillType.FillPatterned
    data: Mapping[str, Any]
    pattern: int 
    _keys: ClassVar[Sequence[str]] = ["Pattern", "ForeColor.RGB", "BackColor.RGB", "Visible"]
    _accessor: ClassVar[CrudeApiAccesssor] = CrudeApiAccesssor(_keys)
    

    @classmethod
    def from_api(cls, api) -> Self:
        data = cls._accessor.read(api)
        return cls(data=data, pattern=api.Pattern)

    def apply_api(self, api):
        api.Patterned(self.pattern)
        self._accessor.write(api, self.data)


class NaiveGradientFillFormat(ApiBridgeBaseModel):
    type: Literal[MsoFillType.FillGradient] = MsoFillType.FillGradient
    data: Mapping[str, Any]

    @classmethod
    def from_api(cls, api) -> Self:
        data = dict()
        data["Visible"] = api.Visible
        data["GradientStyle"] =  api.GradientStyle
        data["GradientColorType"] = api.GradientColorType
        try:
            data["GradientDegree"] = api.GradientDegree
        except com_error:
            data["GradientDegree"] = 0 
        data["GradientVariant"] = api.GradientVariant
        data["ForeColor.RGB"] = api.ForeColor.RGB
        stops = []
        for stop in api.GradientStops:
            elem = dict()
            elem["Color"] = int(stop.Color)
            elem["Position"] = float(stop.Position)
            elem["Transparency"] = float(stop.Transparency)
            stops.append(elem)
        data["GradientStops"] = stops
        return cls(data=data)

    def apply_api(self, api):
        data = dict(self.data)

        # --- 1. PowerPoint の制約の補正 -----------------------------------
        if data["GradientVariant"] == 0:
            data["GradientVariant"] = 1
        if data["GradientStyle"] == MsoFillType.FillMixed:
            data["GradientStyle"] = MsoFillType.FillGradient

        # --- 2. 基本の Gradient 初期化 -----------------------------------
        color_type = data["GradientColorType"]

        if color_type == 1:  # One Color
            api.OneColorGradient(
                data["GradientStyle"],
                data["GradientVariant"],
                data["GradientDegree"],
            )
        elif color_type == 2:  # Two colors
            api.TwoColorGradient(
                data["GradientStyle"],
                data["GradientVariant"],
            )
        else:
            # fallback?
            return api

        # --- 3. GradientStops を全削除（PowerPoint 必須処理） -----------
        # Delete は後ろから行う必要がある
        count = api.GradientStops.Count
        for i in range(count, 0, -1):
            try:
                api.GradientStops.Delete(i)
            except com_error:
                pass

        stops = data["GradientStops"]
        stops_sorted = sorted(stops, key=lambda s: s["Position"])
        for s in stops_sorted:
            # Insert returns a Stop object
            try:
                stop = api.GradientStops.Insert(s["Position"])
            except com_error:
                pass
            else:
                stop.Color = s["Color"]
                stop.Transparency = s["Transparency"]
        return api


class NaiveFallbackFormat(ApiBridgeBaseModel):
    type: int
    data: None = None
    
    def apply_api(self, api):
        print("This FillFormat is out of scope", api.Type)
        #api.Type = self.type

    @classmethod
    def from_api(cls, api) -> Self:
        print("This FillFormat is out of scope", api.Type)
        return cls(type=api.Type)

"""Fill Format.

(2020-04-19) Currently, it is far from perfect.
Only ``Patterned`` / ``Solid`` is handled.
"""
type NaiveTypeFormat = NaiveSolidFillFormat | NaivePatternedFillFormat | NaiveGradientFillFormat | NaiveFallbackFormat

class FillFormatApiBridge(ApiBridgeBaseModel):
    api_data: NaiveTypeFormat
    
    def apply_api(self, api: COMObject):
        self.api_data.apply_api(api)

    @classmethod
    def from_api(cls, api: COMObject) -> Self:
        cls_map = get_discriminator_mapping(NaiveTypeFormat, "type")
        type_ = api.Type
        data = cls_map[type_].from_api(api)
        return cls(api_data=data)
    
class FillApiApplicator:
    @classmethod 
    def apply_bool(cls, api: COMObject, value: bool) -> None:
        api.Visible = constants.msoTrue if value else constants.msoFalse

    @classmethod 
    def apply_any(cls, api: COMObject, value: Any) -> None:
        api.Visible = constants.msoTrue
        color = Color(value)
        api.ForeColor.RGB = color.as_int()
        api.Transparency = 1.0 - color.alpha
        api.Solid()
        
    @classmethod
    def apply(cls, api: COMObject, value: Any) -> None:
        if isinstance(value, bool):
            cls.apply_bool(api, value)
        elif value is None:
            api.Visible = False
        else:
            cls.apply_any(api, value)


class FillFormat:
    """Fill Format.

    (2020-04-19) Currently, it is far from perfect.
    Only ``Patterned`` / ``Solid`` is handled.
    """
    def __init__(self, api: COMObject):
        self._api = api
    
    @property
    def api(self) -> COMObject: 
        return self._api

    @property
    def color(self) -> Color | None:
        rgb_value = self.api.ForeColor.RGB
        # Currently, `ForeColor.RGB` is required.
        if rgb_value:
            return Color(rgb_value)
        else:
            return None
        
    def apply(self, value: bool | ColorLike | Self) -> None:
        if isinstance(value, FillFormat):
            api_bridge = FillFormatApiBridge.from_api(value.api)
            api_bridge.apply_api(self.api)
        else:
            FillApiApplicator.apply(self.api, value)
        
    def __eq__(self, other: Any) -> bool:
        if not isinstance(other, FillFormat):
            return NotImplemented
        return FillFormatApiBridge.from_api(self.api) == FillFormatApiBridge.from_api(other.api)


class FillFormatProperty:
    def __get__(self, shape: "Shape", objtype: type | None = None) -> FillFormat:
        return FillFormat(shape.api.Fill)

    def __set__(self, shape: "Shape", value: bool | FillFormat | ColorLike | None) -> None:
        FillFormat(shape.api.Fill).apply(value)
