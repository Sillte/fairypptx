#from fairypptx.core.resolvers import 
import numpy as np
from typing import Sequence, Self, TYPE_CHECKING, Iterator, cast
from fairypptx.core.types import COMObject, PPTXObjectProtocol
from fairypptx.table.cell import Cell, CellRange
from fairypptx.object_utils import ObjectItems

if TYPE_CHECKING:
    from fairypptx.shape import Shape

class RowColumnMixin(PPTXObjectProtocol):
    """Common Implementation for Row and Column.
    """

    def __getitem__(self, index: int | Sequence[int] | slice) -> Cell | CellRange:
        return CellRange(self.api.Cells)[index]


    def __len__(self):
        return self.api.Cells.Count
    
    def __iter__(self: Self) -> Iterator[Cell]:
        for index in range(len(self)):
            yield cast(Cell, self[index])

    def delete(self) -> None:
        self.api.Delete()

    def select(self) -> None:
        self.api.Select()

    def tighten(self):
        raise NotImplementedError("Shold be implemented in Row or Column.")

    @property
    def cells(self) -> Sequence[Cell]: 
        return list(self)

    @property
    def shapes(self) -> Sequence["Shape"]:
        return [cell.shape for cell in self.cells]

    def is_empty(self) -> bool:
        return all(elem.is_empty() for elem in self.cells)

    def tolist(self):
        """
        Note:
        (2021-03-28) Currently, only `text` is converted.
        """
        return [str(cell.text) for cell in self.cells] 


class Row(RowColumnMixin):
    def __init__(self, api: COMObject) -> None:
        self._api = api

    @property
    def api(self) -> COMObject:
        return self._api

    @property
    def height(self) -> float:
        return self.api.Height

    def tighten(self) -> None:
        # Wondering whether it is sufficiently appropriate.
        # When Row is empty, this fucntion may be wrong.
        self.api.Height = 0  # This is a Hack and not ideal. 

class Column(RowColumnMixin):
    def __init__(self, api: COMObject) -> None:
        self._api = api

    @property
    def api(self) -> COMObject:
        return self._api

    @property
    def width(self):
        return self.api.Width

    def tighten(self) -> None:
        shapes = self.shapes
        
        if not shapes:
             return
        def _get_required_width(shape):
            tf = shape.api.TextFrame
            return tf.MarginLeft + tf.MarginRight + tf.TextRange.BoundWidth
            
        try:
            widths = [_get_required_width(shape) for shape in shapes]
            max_width = max(widths)
            self.api.Width = max(max_width, 0.1) 
            
        except AttributeError as e:
            raise AttributeError(f"Failed to calculate required width for a cell in the column. Inner error: {e}")

class RowsColumnsBase[T: RowColumnMixin](PPTXObjectProtocol):
    def __init__(self, api: COMObject, child_cls: type[T]):
        self._api = api
        self.items = ObjectItems[T](self.api, child_cls)

    @property
    def api(self) -> COMObject:
        return self._api

    def __iter__(self) -> Iterator[T]:
        return iter(self.items)

    def __getitem__(self, key: int | slice | Sequence[int]) -> T | Sequence[T]:
        return self.items[key]
        
    def __len__(self) -> int:
        return len(self.items)

    def delete(self, obj: int | slice | Sequence[int]) -> None:
        obj = self.items.normalize(obj)
        if isinstance(obj, int): 
            item = self.items[obj]
            assert not isinstance(item, Sequence)
            item.delete()
        else:
            # It's necessary to `delete` in decreading order.
            indices = np.sort(obj)[::-1]
            for index in indices:
                item = self[index]
                assert not isinstance(item, Sequence)
                item.delete()

    def insert(self, obj: int | slice | Sequence[int], values: None =None):
        """
        [TODO]
        If values are specified,
        then, values are substituted to the added Rows. 
        """
        assert values is None, "Current Limitation."
        if obj == len(self):
            ret = self.items.cls(self.api.Add())
            if values:
                ret.shape = values
            return ret
        elif isinstance(obj, (int, np.number)):
            # Here, change behavior 
            # when `obj` is positive or negative
            # due to the diffrence of forward and backword.
            obj = int(obj)
            if 0 <= obj:
                obj = self.items.normalize(obj)
                ret = self.items.cls(self.api.Add(obj + 1))
            else:
                obj = self.items.normalize(obj) + 1
                if obj == len(self):
                    ret = self.items.cls(self.api.Add())
                else:
                    ret = self.items.cls(self.api.Add(obj + 1))

            if values:
                ret.shape = values
            return ret
        else:   # indices of Sequence.
            obj = self.items.normalize(obj)
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

    def tighten(self) -> None:
        for item in self.items:
            item.tighten()

    def swap(self, i: int, j: int):
        """Swap the `i` element with `j` element.
        [2020/04/25]
        Firstly, consider only the swap of `texts`.

        Wonder
        --------------------
        It is necessary for `multiple indices`?  
        """

        for elem1, elem2 in zip(self.items[i], self.items[j]):
            elem1.shape.texts, elem2.shape.texts = elem2.shape.texts, elem1.shape.texts

    def is_empty(self):
        return all(elem.is_empty() for elem in self.items)


class Rows(RowsColumnsBase[Row]):
    def __init__(self, api: COMObject) -> None:
        super().__init__(api, Row)

class Columns(RowsColumnsBase[Column]):
    def __init__(self, api: COMObject) -> None:
        super().__init__(api, Column)


