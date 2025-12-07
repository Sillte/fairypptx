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
from fairypptx.object_utils import ObjectDictMixin, to_api2, getattr, is_object, setattr, ObjectDictMixin2
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
        print("arg", arg)
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
            self.font = target.font
            self.paragraphformat = target.paragraphformat


from fairypptx._text.paragraph_format import ParagraphFormat
from fairypptx._text.font import Font


if __name__ == "__main__":
    from fairypptx import TextRange
    from fairypptx import Shape, Shapes, Slide
    #Shape().textrange.register("A"); exit(0)
    #Shape().textrange.like("A"); exit(0)
    # Shape.make("HOGEHOIGE")
    #Shape().textrange.like("TEST"); exit(0)
    #TextRange().register("TEST"); exit(0)
    #$Shape().register("B"); exit(0)
