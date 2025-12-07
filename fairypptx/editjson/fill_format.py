from pydantic import BaseModel, JsonValue, TypeAdapter, Field

from enum import IntEnum
from typing import Any, Mapping, Literal, ClassVar, Sequence, Self, Annotated, cast
from fairypptx._shape import FillFormat
from fairypptx.constants import msoFillSolid, msoFillPatterned, msoFillGradient
from fairypptx.object_utils import setattr, getattr
from fairypptx.editjson.utils import get_discriminator_mapping, CrudeApiAccesssor
from pprint import pprint
from fairypptx.enums import MsoFillType
from fairypptx.editjson.protocols import ApiApplyBaseModel
from pywintypes import com_error



class NaiveSolidFillFormat(ApiApplyBaseModel):
    type: Literal[MsoFillType.FillSolid] = MsoFillType.FillSolid
    data: Mapping[str, Any]
    _keys: ClassVar[Sequence[str]] = ["Type", "ForeColor.RGB", "Visible", "Transparency", "Visible"]
    _accessor: ClassVar[CrudeApiAccesssor] = CrudeApiAccesssor(_keys)

    def apply_api(self, api):
        self._accessor.write(api, self.data)

    @classmethod
    def from_api(cls, api) -> Self:
        data = cls._accessor.read(api)
        return cls.model_validate({"data": data})


class NaivePatternedFillFormat(ApiApplyBaseModel):
    type: Literal[MsoFillType.FillPatterned] = MsoFillType.FillPatterned
    data: Mapping[str, Any]
    _keys: ClassVar[Sequence[str]] = ["Type", "Pattern", "ForeColor.RGB", "BackColor.RGB", "Visible"]
    _accessor: ClassVar[CrudeApiAccesssor] = CrudeApiAccesssor(_keys)
    
    def apply_api(self, api):
        self._accessor.write(api, self.data)

    @classmethod
    def from_api(cls, api) -> Self:
        data = cls._accessor.read(api)
        return cls.model_validate({"data": data})

class NaiveGradientFillFormat(ApiApplyBaseModel):
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
        api.Type = self.type
        # [TODO]: This is not tested nor thoroughly implemented.
        # ## Implementation Policy
        # * TwoColorGradient is always called.  
        # * After that, the settings of Color Stop is modified.
        # [UNSOLVED] Here, `Brightness` of `GradientStop` property exists,
        # However, I cannot get how to read this property. 
        # It seems `self.data["GradientVariant"]` is not used, 
        # When `GradientStyle` maybe ` is ` msoGradientFromCenter`? 
        data = dict(self.data)

        if api.GradientVariable == 0:
            api.GradientVariant = 1
        if data["GradientStyle"] == MsoFillType.FillMixed:
            data["GradientStyle"] = MsoFillType.FillGradient

        api.OneColorGradient(data["GradientStyle"],
                             data["GradientVariant"],
                             data["GradientDegree"],)

        # Currently  `GradientColorType` is not used.
        # api.TwoColorGradient(self.data["GradientStyle"], self.data["GradientVariant"])  

        # Here, Delete is performed, however,
        # 2 colors must be remains. 
        for _ in range(api.GradientStops.Count):
            try:
                api.GradientStops.Delete()
                #print("Delete")
            except com_error:
                pass

        # Copy all the `GradientStop`.  
        for stop in data["GradientStops"]:
            api.GradientStops.Insert2(stop["Color"],
                                      stop["Position"],
                                      Transparency=stop["Transparency"])

        # The remained 2 `GradientStop` are deleted here. 
        for _ in range(2):
            api.GradientStops.Delete(1)
            
        return api


class NaiveFallbackFormat(ApiApplyBaseModel):
    type: int
    
    def apply_api(self, api):
        api.Type = self.type

    @classmethod
    def from_api(cls, api) -> Self:
        return cls(type=api.Type)


type NaiveTypeFormat = NaiveSolidFillFormat | NaivePatternedFillFormat | NaiveGradientFillFormat | NaiveFallbackFormat


class NaiveFillFormatStyle(BaseModel):
    body: NaiveTypeFormat

    @classmethod
    def from_entity(cls, entity: FillFormat) -> Self:
        """Generate itself from the entity of `fairpptx.PPTXObject`
        """
        cls_map = get_discriminator_mapping(NaiveTypeFormat, "type")
        type_ = entity.api.Type
        body = cls_map[type_].from_api(entity.api)
        return cls.model_validate({"body": body})


    def apply(self, entity: FillFormat) -> FillFormat:
        """Apply this edit param to 
        """
        self.body.apply_api(entity.api)
        return entity

if __name__ == "__main__":
    from fairypptx import Shape  
    shape = Shape()
    target = NaiveFillFormatStyle.from_entity(shape.fill)
    print(target.model_dump_json())

