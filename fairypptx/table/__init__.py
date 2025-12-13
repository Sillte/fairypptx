"""table class 

* `DFTable` enables you to get `pandas.DataFrame` from `Table` Object.
* 

Desire:
----------
* Currenly, `Table` / `DFTable` cannot be read from the existing table.
(2021-03-28) Well, then I'd like to attack this problem.

"""

from typing import cast, Sequence, Any
import numpy as np
from fairypptx.apis.table.table_api_writer import TableApiWriter
from fairypptx.apis.table import TableApiModel, TableApiApplicator
from fairypptx.apis.table import CellApiApplicator

from fairypptx.core.resolvers import resolve_table, resolve_shapes
from fairypptx.core.types import COMObject, PPTXObjectProtocol

from fairypptx import registry_utils

from fairypptx.table.cell_containers import Rows, Columns, Row, Column, Cell
from fairypptx.table.cell import Cell
from fairypptx.registry_utils import BaseModelRegistry

class Table:
    def __init__(self, arg: COMObject | PPTXObjectProtocol=None) -> None:
        self._api = resolve_table(arg)

    @property
    def api(self) -> COMObject:
        return self._api

    @property
    def shape(self) -> "Shape":
        from fairypptx.shape import Shape
        return Shape(self.api.Parent)

    @property
    def size(self) -> tuple[int, int]:
        # Naming is under consideration. `row` and `column` are more appropriate?
        return (len(self.api.Rows), len(self.api.Columns))

    @property
    def rows(self) -> Rows:
        return Rows(self.api.Rows)

    @property
    def columns(self) -> Columns:
        return Columns(self.api.Columns)

    def __setitem__(self, key: tuple[int, int], value):
        cell = self[key]
        assert isinstance(cell, Cell)
        cell.shape.text = str(value)


    def __getitem__(self, key: tuple[int, int] | tuple[slice, int] | tuple[int | slice] | tuple[slice | slice]) -> Cell | Sequence[Cell] | Sequence[Sequence[Cell]]:
        if isinstance(key, tuple):
            r_size, c_size = self.size
            if len(key) == 2:
                if isinstance(key[0], int) and isinstance(key[1], int):
                    i_row, i_column = key
                    cell_object = self.api.Cell(i_row + 1, i_column + 1)
                    return Cell(cell_object)
                elif isinstance(key[0], slice) and isinstance(key[1], int):
                    column = self.columns[key[1]]
                    return column[key[0]]
                elif isinstance(key[0], int) and isinstance(key[1], slice):
                    row = self.rows[key[0]]
                    return row[key[1]]
                elif isinstance(key[0], slice) and isinstance(key[1], slice):
                    r_indices, c_indices = range(*key[0].indices(r_size)), range(*key[1].indices(c_size))
                    return [[Cell(self.api.Cell(r_index + 1, c_index + 1)) for c_index in c_indices] for r_index in r_indices] 

        raise ValueError(f"`{key=}` cannot be interpreted.")

    @staticmethod
    def make(arg=None, size:tuple[int, int] | None = None, **kwargs) -> "Table":
        return TableFactory.make_table(arg, size=size, **kwargs)

    @staticmethod
    def empty(shape: tuple[int, int]) -> "Table":
        return TableFactory.empty(*shape)

    def tighten(self):
        for row in self.rows:
            row.tighten()
        for column in self.columns:
            column.tighten()

    def tolist(self):
        data = [[str(cell.text) for cell in row] for row in self.rows]
        return data

    def to_numpy(self):
        """Convert to `numpy`.
        """
        data = self.tolist()
        return np.array(data)

    @property
    def values(self):
        return self.to_numpy()  

    def register(self, style: str, style_type: str | None | type = None) -> None:
        from fairypptx.editjson.style_type_registry import TableStyleTypeRegistry
        if not isinstance(style_type, type): 
            style_type = TableStyleTypeRegistry.fetch(style_type)
        edit_param = style_type.from_entity(self)
        BaseModelRegistry.put(edit_param, "Table", style)


    def like(self, style: str):
        from fairypptx.editjson.protocols import EditParamProtocol
        edit_param = BaseModelRegistry.fetch("Table", style)
        edit_param = cast(EditParamProtocol, edit_param)
        edit_param.apply(self)


class TableFactory:

    @staticmethod
    def make_table(arg=None, size:tuple[int, int] | None = None, **kwargs) -> Table:
        if arg is None:
            if size: 
                return TableFactory.empty(*size)
            else:
                return TableFactory.empty()

        if isinstance(arg, tuple) and len(arg) == 2 and isinstance(arg[0], int) and isinstance(arg[0], int):
            return TableFactory.empty(*arg)

        if isinstance(arg, np.ndarray):
            return TableFactory.from_numpy(arg)
        raise ValueError(arg)


    @staticmethod
    def from_numpy(arr: np.ndarray) -> Table:
        arr = np.atleast_2d(arr)
        n_row, n_col = arr.shape

        table = TableFactory.empty(n_row, n_col)
        for r in range(n_row):
            for c in range(n_col):
                table[r, c] = arr[r, c]
        return table

    @staticmethod
    def empty(n_row: int = 3, n_col: int = 2) -> Table:
        shapes_api = resolve_shapes()
        shape_api = shapes_api.AddTable(NumRows=n_row, NumColumns=n_col)
        return Table(shape_api.Table)

class TableProperty:
    def __get__(self, parent: PPTXObjectProtocol, objtype=None):
        return Table(parent.api.Table)

    def __set__(self, parent: PPTXObjectProtocol, value: Any) -> None:
        TableApiApplicator.apply(parent.api.Table, value)

        

if __name__ == "__main__":
    table = Table()
    array = np.zeros(shape=(4, 2)).astype(object)
    table[0:1, :2] = "TARET"
    array = np.array([[1, 2], [3, 4]])
    print(slice(0, 2))
    array[0:2, 1] = [0.5, 0.6]
    print(array)
