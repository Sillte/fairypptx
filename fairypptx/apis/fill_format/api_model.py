"""Fill Format API Schema.

FillFormat API holds different attribute sets depending on Fill Type 
(Solid/Patterned/Gradient). This module uses tagged union (discriminated union)
to represent these variants clearly:

- NaiveSolidFillFormat: RGB, Transparency, Visible
- NaivePatternedFillFormat: Pattern, ForeColor, BackColor, Visible  
- NaiveGradientFillFormat: GradientStyle, GradientStops, GradientVariant, ...
- NaiveFallbackFormat: Fallback for unsupported fill types

Pydantic discriminator (via 'type' field) ensures that from_api() correctly 
selects the appropriate variant based on api.Type. This provides type safety 
and clear schema validation.

Design rationale:
  - A single FillFormatApiModel(type, data) would obscure which keys are valid
  - Tagged union makes the schema explicit and Pydantic-compatible
  - Each variant documents its own responsibility: what COMObject keys to read/write
"""

from pywintypes import com_error
from fairypptx.core.models import BaseApiModel
from fairypptx.core.types import COMObject
from fairypptx.core.utils import CrudeApiAccesssor, get_discriminator_mapping
from fairypptx.enums import MsoFillType


from typing import Any, ClassVar, Literal, Mapping, Self, Sequence


class NaiveSolidFillFormat(BaseApiModel):
    type: Literal[MsoFillType.FillSolid] = MsoFillType.FillSolid
    data: Mapping[str, Any]
    _keys: ClassVar[Sequence[str]] = ["ForeColor.RGB", "Visible", "Transparency"]
    _accessor: ClassVar[CrudeApiAccesssor] = CrudeApiAccesssor(_keys)

    def apply_api(self, api):
        self._accessor.write(api, self.data)

    @classmethod
    def from_api(cls, api) -> Self:
        data = cls._accessor.read(api)
        return cls(data=data)


class NaivePatternedFillFormat(BaseApiModel):
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


class NaiveGradientFillFormat(BaseApiModel):
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


class NaiveFallbackFormat(BaseApiModel):
    type: int
    data: None = None

    def apply_api(self, api):
        print("This FillFormat is out of scope", api.Type)
        #api.Type = self.type

    @classmethod
    def from_api(cls, api) -> Self:
        print("This FillFormat is out of scope", api.Type)
        return cls(type=api.Type)


type NaiveTypeFormat = NaiveSolidFillFormat | NaivePatternedFillFormat | NaiveGradientFillFormat | NaiveFallbackFormat


class FillFormatApiModel(BaseApiModel):
    api_data: NaiveTypeFormat

    def apply_api(self, api: COMObject):
        self.api_data.apply_api(api)

    @classmethod
    def from_api(cls, api: COMObject) -> Self:
        cls_map = get_discriminator_mapping(NaiveTypeFormat, "type")
        type_ = api.Type
        data = cls_map[type_].from_api(api)
        return cls(api_data=data)
