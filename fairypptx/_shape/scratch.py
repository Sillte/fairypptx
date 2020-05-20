from fairypptx._shape import FillFormat, LineFormat
from fairypptx import Shapes, Shape

fill = FillFormat()
a = Shape()
fill["Visible"] = False
a.fill = fill
print(a.fill)

pass
