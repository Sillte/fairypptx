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
from fairypptx.object_utils import ObjectDictMixin
from fairypptx import registory_utils

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
            self.font = registory_utils.fetch("Font", target)
            self.paragraphformat = registory_utils.fetch("ParagraphFormat", target)
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
    def color(self):
        return Color(self.data["Color.RGB"])

    @color.setter
    def color(self, value):
        self["Color.RGB"] = Color(value).as_int()

    @property
    def size(self):
        return self["Size"]

    @size.setter
    def size(self, value):
        self["Size"] = value


class ParagraphFormat(ObjectDictMixin):
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
    data["FarEastLineBreakControl"] = constants.msoFalse
    data["Alignment"] = constants.ppAlignLeft
    data["BaseLineAlignment"] = constants.ppBaselineAlignBaseline

    data["HangingPunctuation"] = constants.msoFalse

    data["LineRuleAfter"] = 0
    data["LineRuleBefore"] = 0
    data["LineRuleWithin"] = constants.msoTrue
    data["SpaceAfter"] = 0
    data["SpaceBefore"] = 0
    data["SpaceWithin"] = constants.msoFalse

    data["Bullet.Visible"] = constants.msoFalse
    data["Bullet.Character"] = 8226
    data["Bullet.Font.Name"] = "Arial"
    data["Bullet.Type"] = constants.ppBulletUnnumbered



if __name__ == "__main__":
    from fairypptx import TextRange
    f1 = Font()
    tr = TextRange()
    p = ParagraphFormat(tr.ParagraphFormat)
    print(p)
    s = tr.font
    f2 = Font(tr.Font)

