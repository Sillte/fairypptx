from pywintypes import com_error

from collections import defaultdict
from fairypptx import constants
from fairypptx.color import Color
from fairypptx.object_utils import ObjectDictMixin, getattr


class FillFormat(ObjectDictMixin):
    """Fill Format.

    (2020-04-19) Currently, it is far from perfect.
    Only ``Patterned`` / ``Solid`` is handled.
    """

    data = dict()
    data["Type"] = constants.msoFillSolid
    data["ForeColor.RGB"] = 0
    data["Visible"] = constants.msoFalse
    data["Transparency"] = 0

    readonly = ["Type", "Pattern", "GradientStyle", "GradientColorType", "GradientVariant", "GradientStops", "GradientDegree"]  # readonly parameters.

    common_keys = ["Type", "Visible"]
    type_to_keys = defaultdict(list)
    type_to_keys[constants.msoFillSolid] = ["ForeColor.RGB", "Visible", "Transparency"]
    type_to_keys[constants.msoFillPatterned] = [
        "Pattern",
        "ForeColor.RGB",
        "BackColor.RGB",
    ]
    # For constants.msoFillGradient, the function is required..

    def __init__(self, arg=None, **kwargs):
        super().__init__(arg, **kwargs)
        assert "Type" in self.data

    def to_dict(self, api_object):
        data = {key: getattr(api_object, key) for key in self.common_keys}
        # data["Type"] = type_valkue
        type_value = getattr(api_object, "Type")
        if type_value in self.type_to_keys:
            return dict(data,  **{key: getattr(api_object, key) for key in self.type_to_keys[type_value]})
        elif type_value == constants.msoFillGradient:
            data["GradientStyle"] =  api_object.GradientStyle
            data["GradientColorType"] = api_object.GradientColorType
            try:
                data["GradientDegree"] = api_object.GradientDegree
            except com_error:
                data["GradientDegree"] = 0 
                pass
            data["GradientVariant"] = api_object.GradientVariant
            data["ForeColor.RGB"] = api_object.ForeColor.RGB
            stops = []
            for stop in api_object.GradientStops:
                elem = dict()
                elem["Color"] = int(stop.Color)
                elem["Position"] = float(stop.Position)
                elem["Transparency"] = float(stop.Transparency)
                stops.append(elem)
            data["GradientStops"] = stops
            return data
        return data

    def apply(self, api_object):
        type_value = self.data["Type"]
        if type_value == constants.msoFillSolid:
            api_object.Solid()
        elif type_value == constants.msoFillPatterned:
            api_object.Patterned(self.data["Pattern"])
        elif type_value == constants.msoFillGradient:
            g_style = self.data["GradientStyle"]

            # ## Implementation Policy
            # * TwoColorGradient is always called.  
            # * After that, the settings of Color Stop is modified.
            # [UNSOLVED] Here, `Brightness` of `GradientStop` property exists,
            # However, I cannot get how to read this property. 
            
            # It seems `self.data["GradientVariant"]` is not used, 
            # When `GradientStyle` maybe ` is ` msoGradientFromCenter`? 
            if self.data["GradientVariant"] == 0:
                self.data["GradientVariant"] = 1

            # This is a last resort and it is not good.
            if self.data["GradientStyle"] == constants.msoGradientMixed:
                print("GradientStyle is Mixed, so this cannot be handled correctly.")
                self.data["GradientStyle"] = constants.msoGradientHorizontal

            api_object.OneColorGradient(self.data["GradientStyle"],
                                        self.data["GradientVariant"],
                                        self.data["GradientDegree"],)
            # Currently  `GradientColorType` is not used.
            # api_object.TwoColorGradient(self.data["GradientStyle"], self.data["GradientVariant"])  

            # Here, Delete is performed, however,
            # 2 colors must be remains. 
            for _ in range(api_object.GradientStops.Count):
                try:
                    api_object.GradientStops.Delete(); print("Delete")
                except com_error as e:
                    pass
            assert api_object.GradientStops.Count == 2

            # Copy all the `GradientStop`.  
            for stop in self.data["GradientStops"]:
                api_object.GradientStops.Insert2(stop["Color"],
                                                 stop["Position"],
                                                 Transparency=stop["Transparency"])

            # The remained 2 `GradientStop` are deleted here. 
            for _ in range(2):
                api_object.GradientStops.Delete(1)
        else:
            raise ValueError(f"Currently `type_value`={type_value} cannot be handled.")

        super().apply(api_object)

    def __eq__(self, other):
        """Unrelated keys are excluded for comparison.  
        Nevertheless, I think comparision is not complete. 
        """
        untargets = set(["GradientColorType", "GradientDegree"])
        left = {key: value for key, value in self.items() if key not in untargets}
        right = {key: value for key, value in other.items() if key not in untargets}
        return left == right
        #flag = (left.keys() == right.keys())
        #for key in left.keys() | right.keys():
        #    if left.get(key, None) != right.get(key, None):
        #        print("diff", left.get(key), right.get(key), key)
        #print("flag,", flag)
        #return flag


    @property
    def color(self):
        rgb_value = self.get("ForeColor.RGB", None)
        # Currently, `ForeColor.RGB` is required.
        if rgb_value:
            return Color(rgb_value)
        else:
            return None

class FillFormatProperty:
    def __get__(self, shape, objtype=None):
        return FillFormat(shape.api.Fill)

    def __set__(self, shape, value):
        Fill = shape.api.Fill
        if value is None:
            Fill.Visible = constants.msoFalse
        elif isinstance(value, FillFormat):
            value.apply(Fill)
        else:
            Fill.Visible = constants.msoTrue
            color = Color(value)
            Fill.ForeColor.RGB = color.as_int()
            Fill.Transparency = 1.0 - color.alpha
            Fill.Solid()
