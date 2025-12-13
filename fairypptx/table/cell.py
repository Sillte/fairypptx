from typing import Sequence, Iterator
from collections.abc import Sequence as SeqABC 

from fairypptx.core.types import COMObject, PPTXObjectProtocol
from fairypptx.apis.table.api_model import CellApiModel
from fairypptx.apis.table.applicator import CellApiApplicator
from fairypptx import object_utils
from fairypptx.object_utils import is_object


class Cell:
    def __init__(self, api:COMObject | PPTXObjectProtocol) -> None:
        if isinstance(api, PPTXObjectProtocol):
            api = api.api
        self._api = api
        
    @property
    def api(self) -> COMObject:
        return self._api

    @property
    def shape(self) -> "Shape":
        from fairypptx.shape import Shape
        return Shape(self.api.Shape)

    @property
    def text(self):
        return self.shape.text

    def is_empty(self) -> bool:
        return self.shape.text.strip() == ""


class CellRange:
    """1DArray of `Cell`, mainly it is expected to be called from `Row` or `Column`, not `Table` directly.
    """
    def __init__(self, arg:COMObject | PPTXObjectProtocol | Sequence[PPTXObjectProtocol]) -> None:
        self._cells: list[Cell] = self._solve_cells(arg)

    def __len__(self) -> int:
        return len(self._cells)

    def __iter__(self) -> Iterator[Cell]:
        yield from self._cells

    def __getitem__(self, key: int | Sequence[int] | slice) -> "Cell | CellRange":
        if isinstance(key, int):
            return self._cells[key]
        elif isinstance(key, slice):
            return CellRange(self._cells[key])
        elif isinstance(key, Sequence):
            return CellRange([self._cells[elem] for elem in key])
        else:
            raise TypeError(f"Invalid key: {key!r}")

    @property
    def api(self) -> COMObject:
        return self._cells[0].api.Parent.CellRange


    def _solve_cells(self, arg) -> list[Cell]:
        """Normalize input → Sequence[Cell]"""

        if is_object(arg, "CellRange"):
            return [Cell(arg.Item(i + 1)) for i in range(arg.Count)]

        # 2) Python list of Shape
        if isinstance(arg, SeqABC) and not isinstance(arg, (str, bytes)):
            if all(isinstance(s, Cell) for s in arg):
                return list(arg)

        if isinstance(arg, CellRange):
            return list(arg._cells)
        raise ValueError(f"`{arg=}` cannot be intepretted in _`_solve_cells.`")

