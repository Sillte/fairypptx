from _ctypes import COMError
from collections.abc import Sequence
from PIL import Image
import numpy as np
import pandas as pd
from fairypptx import constants
from fairypptx.constants import msoTrue, msoFalse

from fairypptx.color import Color
from fairypptx.box import Box, intersection_over_cover
from fairypptx.application import Application
from fairypptx.slide import Slide
from fairypptx.shape import  Shapes, Shape
from fairypptx.object_utils import is_object, upstream, stored, ObjectClassMixin, ObjectItems

class Cell(ObjectClassMixin):
    @property
    def shape(self):
        return Shape(self.api.Shape)

    @property
    def text(self):
        return self.shape.text

    def is_empty(self):
        """ 
        (2020-04-25): Currently, only text is handled.
        """
        return (self.shape.text == "")



class RowColumnMixin(ObjectClassMixin):
    """Common Implementation for Row and Column
    """
    def __getitem__(self, index):
        if isinstance(index, int):
            return Cell(self.api.Cells.Item(index + 1))
        raise TypeError(f"Type of index is invalid; ")

    def __len__(self):
        return self.api.Cells.Count
    
    def __iter__(self):
        for index in range(len(self)):
            yield self[index]

    def delete(self):
        self.api.Delete()

    def select(self):
        self.api.Select()

    def tighten(self):
        raise NotImplementedError("Shold be implemented in Row or Column.")

    @property
    def cells(self): 
        return list(self)

    @property
    def shapes(self):
        return [cell.shape for cell in self.cells]

    def is_empty(self):
        return all(elem.is_empty() for elem in self.cells)

class Row(RowColumnMixin):

    @property
    def height(self):
        return self.api.Height

    def tighten(self):
        # Wonder
        # When Row is empty, this fucntion may be wrong.
        self.api.Height = 0  # This is a Hack and not ideal. 

class Column(RowColumnMixin):

    @property
    def width(self):
        return self.api.Width

    def tighten(self):
        shapes = self.shapes
        def _to_width(shape):
            return shape.api.TextFrame.MarginLeft + shape.api.TextFrame.MarginRight + shape.api.TextFrame.TextRange.BoundWidth
        widths = [_to_width(shape) for shape in shapes]
        self.api.Width = max(widths)


class RowsColumnsMixin(ObjectClassMixin):
    """Common Implementation for `Rows` and `Columns`.

    Note
    ------------------------------
    * Please specify `child_class` to be `Row` or  `Column`.
    """
    # Please specify `child_class`
    child_class = None

    def __init__(self, arg=None):
        super().__init__(arg)
        self.items = ObjectItems(self.api, self.child_class)

    def __getitem__(self, key):
        return self.items[key]
        
    def __len__(self):
        return len(self.items)

    def delete(self, obj):
        obj = self.items.normalize(obj)
        if isinstance(obj, int): 
            self[obj].delete()
        else:
            # It's necessary to `delete` in decreading order.
            indices = np.sort(obj)[::-1]
            for index in indices:
                self[index].delete()

    def insert(self, obj, values=None):
        """
        [TODO]
        If values are specified,
        then, values are substituted to the added Rows. 
        """
        assert values is None, "Current Limitation."
        obj = self.items.normalize(obj)
        if isinstance(obj, int):
            ret = self.items.cls(self.api.Add(obj + 1))
            if values:
                ret.shape = values
            return ret
        else:
            # It's necessary to `insert` in decreading order.
            indices = np.sort(obj)[::-1]
            orders = np.argsort(obj)[::-1]
            targets = [self.items.cls(self.api.Add(index + 1)) for index in indices]
            # Mapping the content.
            ret = [None] * len(indices)
            for order, target in zip(orders, targets):
                ret[order] = target
            if values:
                assert len(ret) == len(values)
                for elem, value in zip(ret, values):
                    elem.shape = value
            return ret

    def tighten(self):
        for item in self.items:
            item.tighten()

    def swap(self, i, j):
        """Swap the `i` element with `j` element.
        [2020/04/25]
        Firstly, consider only the swap of `texts`.

        Wonder
        --------------------
        It is necessary for `multiple indices`?  
        """
        lhs = self.items[i]
        rhs = self.items[j]
        for elem1, elem2 in zip(self.items[i], self.items[j]):
            elem1.shape.texts, elem2.shape.texts = elem2.shape.texts, elem1.shape.texts

    def is_empty(self):
        return all(elem.is_empty() for elem in self.items)


class Rows(RowsColumnsMixin):
    child_class = Row

class Columns(RowsColumnsMixin):
    child_class = Column
    

