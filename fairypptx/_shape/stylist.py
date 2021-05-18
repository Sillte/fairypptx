"""
Note
----

Protocol of `Stylist`.
-----------------------

```
def __init__(self, instance):
    pass

def __call__(self, instance):
    pass
```
* must be able to be saved with `pickle`.

"""

from fairypptx import constants 
from fairypptx._text.textrange_stylist import ParagraphTextRangeStylist
from fairypptx._shape import FillFormat
from fairypptx._shape import LineFormat

class ShapeStylist:
    def __init__(self, shape):
        if shape.text:
            self.text_stylist = ParagraphTextRangeStylist(shape.textrange)
        else:
            self.text_stylist = None
        self.fill = FillFormat(shape.fill)
        self.line = LineFormat(shape.line)
        self.auto_shape_type = shape.AutoShapeType

    def __call__(self, shape):
        if self.text_stylist:
            self.text_stylist(shape.textrange)
        shape.fill = self.fill
        shape.line = self.line
        shape.api.AutoShapeType = self.auto_shape_type
        return shape

