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

class ShapeStylist:
    def __init__(self, shape):
        self.text_stylist = ParagraphTextRangeStylist(shape.textrange)
        self.fill = FillFormat(shape.fill)

    def __call__(self, shape):
        self.text_stylist(shape.textrange)
        shape.fill = self.fill
        return shape

