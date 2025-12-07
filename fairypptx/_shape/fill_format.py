from pywintypes import com_error

from collections import defaultdict
from fairypptx import constants
from fairypptx.color import Color

from typing import Any, Mapping, Literal, ClassVar, Sequence, Self, Annotated, cast
from fairypptx.core.utils import get_discriminator_mapping
from fairypptx.core.utils import CrudeApiAccesssor
from pprint import pprint
from fairypptx.enums import MsoFillType
from fairypptx.core.models import ApiBridgeBaseModel
from pywintypes import com_error


class NaiveSolidFillFormat(ApiBridgeBaseModel):
    type: Literal[MsoFillType.FillSolid] = MsoFillType.FillSolid
    data: Mapping[str, Any]
    _keys: ClassVar[Sequence[str]] = ["ForeColor.RGB", "Visible", "Transparency", "Visible"]
    _accessor: ClassVar[CrudeApiAccesssor] = CrudeApiAccesssor(_keys)

    def apply_api(self, api):
        self._accessor.write(api, self.data)

    @classmethod
    def from_api(cls, api) -> Self:
        api.Solid()
        data = cls._accessor.read(api)
        return cls.model_validate({"data": data})


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
    
    def apply_api(self, api):
        self.api_data.apply_api(api)

    @classmethod
    def from_api(cls, api) -> Self:
        cls_map = get_discriminator_mapping(NaiveTypeFormat, "type")
        type_ = api.Type
        data = cls_map[type_].from_api(api)
        return cls(api_data=data)


class FillFormat:
    """Fill Format.

    (2020-04-19) Currently, it is far from perfect.
    Only ``Patterned`` / ``Solid`` is handled.
    """
    def __init__(self, api):
        self._api = api
    
    @property
    def api(self): 
        return self._api

    @property
    def color(self):
        rgb_value = self.api.ForeColor.RGB
        # Currently, `ForeColor.RGB` is required.
        if rgb_value:
            return Color(rgb_value)
        else:
            return None

    def __eq__(self, other):
        api_bridge = FillFormatApiBridge.from_api(self.api)
        api_bridge1 = FillFormatApiBridge.from_api(other.api)
        return api_bridge.model_dump(exclude_defaults=True)  == api_bridge1.model_dump(exclude_defaults=True)

class FillFormatProperty:
    def __get__(self, shape, objtype=None):
        return FillFormat(shape.api.Fill)

    def __set__(self, shape, value):
        Fill = shape.api.Fill
        if value is None:
            Fill.Visible = constants.msoFalse
        elif isinstance(value, FillFormat):
            api_bridge = FillFormatApiBridge.from_api(value.api)
            api_bridge.apply_api(shape.api.Fill)
        else:
            Fill.Visible = constants.msoTrue
            color = Color(value)
            Fill.ForeColor.RGB = color.as_int()
            Fill.Transparency = 1.0 - color.alpha
            Fill.Solid()