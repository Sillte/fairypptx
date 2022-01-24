from fairypptx.shape import Shapes, Shape
from fairypptx.shape import Shapes
import _ctypes 
from pywintypes import com_error

from fairypptx.slide import Slide


class ShapesSelector:
    def __init__(self, target=None):
        if target is None:
            target = Slide()
        self.target = target

    def match(self, query) -> Shape:
        """Return one shape, satisfying the condition of `query`.
        """
        if callable(query):
            for shape in self._to_shapes(): 
                try:
                    if func(shape):
                        return shape
                except com_error:
                    pass
            return None
        if isinstance(query, str):
            query = query.strip()
            for shape in self._to_shapes():
                try:
                    if shape.text.strip().startswith(query):
                        return shape
                except com_error:
                    pass
            return None
        return None

    def findall(self, query):
        result = []
        if callable(query):
            for shape in self._to_shapes(): 
                try:
                    if func(shape):
                        result.append(shape)
                except _ctypes.COMError:
                    pass
        elif isinstance(query, str):
            for shape in self._to_shapes(): 
                try:
                    if shape.text.strip().find(query) != -1:
                        result.append(shape)
                except com_error:
                    pass
        if result:
            return Shapes(result)
        else:
            return []

    def _to_shapes(self):
        if isinstance(self.target, Slide):
            for shape in Slide().leaf_shapes:
                yield shape
            return
