""" An test class for Text Format.


Note
---------------------------------------

`data` of class is a template.


Wonder
-----------------------------------
For Font and ParagraphFormat, what kind of Mapping can be accepted? 
If we want restriction, for example, ...
```
assert Font.data.keys().issuperclass(arg) # or
assert Font.data.keys().issubclass(arg) # ...
```
No restriction is one strategy.

"""
from collections import UserDict, UserString
from collections.abc import Mapping 

from fairypptx import constants
from fairypptx.color import Color
from fairypptx import object_utils
from fairypptx.object_utils import ObjectDictMixin, to_api2, getattr, is_object, setattr
from fairypptx import registry_utils

class Text(UserString):
    """Represents the format of TextRange.  

    This is a subclass of UserString.
    Hence, this class behaves as `str`.
    Not only the content of `str`,  it includes the information about Format and Indent.

    Note
    -----------------
    Attributes:
        data (str): The content of str.
        font (UserDict): Represents the information of `Font`
        paragraphformat(UserDict): Represents the information of `ParagraphFormat`.
    """

    def __init__(self, arg, **kwargs):
        self.data = self._construct_data(arg)
        self.font, self.paragraphformat = self._construct_format(arg)

    def _construct_data(self, arg):
        from fairypptx import TextRange
        if isinstance(arg, UserString): 
            return str(arg)
        elif isinstance(arg, TextRange):
            return str(arg.api.Text)
        elif object_utils.is_object(arg, "TextRange"):
            return arg.Text
        elif isinstance(arg, str):
            return arg
        raise ValueError(f"`{type(arg)}`, `{arg}` is an invalid parameter.")

    def _construct_format(self, arg):
        from fairypptx import TextRange
        if isinstance(arg, Text):
            return arg.font, arg.paragraphformat 
        elif isinstance(arg, TextRange):
            return arg.font, arg.paragraphformat  
        elif object_utils.is_object(arg, "TextRange"):
            return Font(arg.Font), ParagraphFormat(arg.ParagraphFormat)
        return Font({}), ParagraphFormat({})

    def register(self, key, disk=False):
        self.font.register(key, disk=disk)
        self.paragraphformat.register(key, disk=disk)

    
    def like(self, target):
        if isinstance(target, str):
            self.font = registry_utils.fetch("Font", target)
            self.paragraphformat = registry_utils.fetch("ParagraphFormat", target)
        elif isinstance(target, (Text, TextRange)):
            self.font = dict(target.font)
            self.paragraphformat = dict(target.paragraphformat)


class Font(ObjectDictMixin):
    """Represents the Font Information. 
    """
    data = dict()
    data["Size"] = 18
    data["Name"] = ""
    data["Bold"] = constants.msoFalse
    data["Italic"] = constants.msoFalse
    data["Shadow"] = constants.msoFalse
    data["Superscript"] = constants.msoFalse
    data["Subscript"] = constants.msoFalse
    data["Underline"] = constants.msoFalse
    data["Color.RGB"] = 0

    @property
    def bold(self):
        return self["Bold"] != constants.msoFalse

    @bold.setter
    def bold(self, value):
        if value is True:
            self["Bold"] = constants.msoTrue
        elif value is False:
            self["Bold"] = constants.msoFalse
        else:
            self["Bold"] = value

    @property
    def underline(self):
        return self["Underline"] != constants.msoFalse

    @bold.setter
    def underline(self, value):
        if value is True:
            self["Underline"] = constants.msoTrue
        elif value is False:
            self["Underline"] = constants.msoFalse
        else:
            self["Underline"] = value

    @property
    def color(self):
        return Color(self.data["Color.RGB"])

    @color.setter
    def color(self, value):
        value = Color(value)
        self["Color.RGB"] = value.as_int()
        if value.alpha < 1:
            api2 = to_api2(self.api)
            api2.Fill.Transparency = 1 - value.alpha 

    @property
    def size(self):
        return self["Size"]

    @size.setter
    def size(self, value):
        self["Size"] = value


class ParagraphFormat(Mapping):
    """Represents the Font Information. 

    Note
    -------------------------------------
    Curently, About `data`, the order of key is important
    since some keys (I infer ``Bullet`.Character`?) change the other properties implicitly.
    This knowledge must be also taken care by users to customize.
    [TODO] You can modify this. See ``FillFormat``.


    Wonder
    -----------------------------------------
    BulletFormat is introduced or not.
    * https://docs.microsoft.com/ja-jp/office/vba/api/powerpoint.bulletformat.number
    When there is a tree structure of ObjectDictMixin exist, `apply` should be modified.

    """
    data = dict()
    data2 = dict()

    data["FarEastLineBreakControl"] = constants.msoFalse
    data["Alignment"] = constants.ppAlignLeft
    data["BaseLineAlignment"] = constants.ppBaselineAlignBaseline
    data["HangingPunctuation"] = constants.msoFalse
    data["LineRuleAfter"] = None
    data["LineRuleBefore"] = None
    data["LineRuleWithin"] = None
    data["SpaceAfter"] = None
    data["SpaceBefore"] = None
    data["SpaceWithin"] = constants.msoFalse

    data["Bullet.Visible"] = None
    data["Bullet.Character"] = None
    data["Bullet.Font.Name"] = None
    data["Bullet.Type"] = None

    data2["FirstLineIndent"] = None
    data2["LeftIndent"] = None

    readonly = []
    name = "ParagraphFormat"

    def __init__(self, arg=None):
        self._construct(arg)

    def _construct(self, arg):
        cls = type(self)
        self._api = None
        if arg is None:
            self.data = cls.data.copy()
            self.data2 = cls.data2.copy()
        elif is_object(arg, cls.name):
            self._api = arg
            self.data = {key: getattr(self.api, key) for key in cls.data}
            self.data2 = {key: getattr(self.api2, key) for key in cls.data2}
        elif isinstance(arg, Mapping):
            for key, value in arg.items(): 
                self[key] = value
        else:
            raise ValueError("Given `arg` is not appropriate.", arg)

    @property
    def api(self):
        return self._api

    @property
    def api2(self):
        if self.api:
            return to_api2(self.api)
        else:
            return None

    def __repr__(self):
        return repr(self.data) + "\n" + repr(self.data2)

    def items(self):
        import itertools 
        return itertools.chain(self.data.items(), self.data2.items())

    def __len__(self):
        return len(self.data) + len(self.data2)

    def __getitem__(self, key):
        if key in self.data:
            return self.data[key]
        if key in self.data2:
            return self.data2[key]
        raise KeyError(key)

    def __setitem__(self, key, item):
        if key in self.data:
            if self.api:
                if item is not None:
                    setattr(self.api, key, item)
                self.data[key] = item
            return 
        elif key in self.data2:
            if self.api:
                if item is not None:
                    setattr(self.api2, key, item)
            self.data2[key] = item
            return
        elif self.api: 
            # This is a fallback, 
            try:
                setattr(self.api, key, item)
            except AttributeError as e:
                pass
            else:
                self.data[key] = item
                return 
            try:
                setattr(self.api2, key, item)
            except AttributeError as e:
                pass
            else:
                self.data2[key] = item
                return 
        raise KeyError("Cannot set key", key)

    def __delitem__(self, key):
        del self.data[key]
        del self.data2[key]

    def __iter__(self):
        import itertools
        return itertools.chain(iter(self.data), iter(self.data2))

    def __contains__(self, key):
        return key in self.data or key in self.data2

    def __getstate__(self):
        """For `pickle` serialization
        """
        return {"name": self.name,
                "data": self.data,
                "data2": self.data2,
                "readonly": self.readonly}

    def register(self, key, disk=False):
        """Register to the storage."""
        name = type(self).name
        registry_utils.register(name, key, self, extension=".pkl", disk=disk)


    @classmethod
    def fetch(cls, key, disk=True):
        """Construct the instance with `key` object."""
        name = cls.name
        return registry_utils.fetch(name, key, disk=True)


if __name__ == "__main__":
    from fairypptx import TextRange
    from fairypptx import Shape, Shapes, Slide
    # Shape.make("HOGEHOIGE")
    Shape().textrange.like("TEST"); exit(0)
    #TextRange().register("TEST"); exit(0)
    shapes = Slide().shapes
    f = shapes[0].textrange.paragraphs[0].paragraphformat
    f["Application"] = None
    print(set(f.values()))
    print(dict(f)); exit(0)
    f.register("TEST")
    m = ParagraphFormat.fetch("TEST")
    print(m)
    shapes[1].textrange.paragraphformat = f

    
    f1 = Font()
    tr = Shape().textrange
    p = ParagraphFormat(tr.ParagraphFormat)
    print(p)
    import pickle
    s = tr.font
    f2 = Font(tr.Font)

