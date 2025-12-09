"""table class 

* `DFTable` enables you to get `pandas.DataFrame` from `Table` Object.
* 

Desire:
----------
* Currenly, `Table` / `DFTable` cannot be read from the existing table.
(2021-03-28) Well, then I'd like to attack this problem.

"""

import numpy as np
from fairypptx._table.table_api_writer import TableApiWriter


from fairypptx.core.resolvers import resolve_table, resolve_shapes
from fairypptx.core.types import COMObject

from fairypptx import registry_utils

from fairypptx._table import Cell, Rows, Columns, Row, Column
from fairypptx._table.table_api_writer import TableApiWriter 

class Table:
    def __init__(self, arg=None):
        self._api = resolve_table(arg)

    @property
    def shape(self) -> "Shape":
        from fairypptx.shape import Shape
        return Shape(self.api.Parent)

    @property
    def api(self) -> COMObject:
        return self._api

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

    def __setitem__(self, key, value):
        writer = TableApiWriter(self.api)
        writer.write(key, value)

    def __getitem__(self, key):
        i_row, i_column = key
        cell_object = self.api.Cell(i_row + 1, i_column + 1)
        return Cell(cell_object)

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
        data = [[str(cell.text) for cell in row.cells] for row in self.rows]
        return data

    def to_numpy(self):
        """Convert to `numpy`.
        """
        data = self.tolist()
        return np.array(data)

    @property
    def values(self):
        return self.to_numpy()  

    def register(self, key, disk=True):
        from fairypptx.editjson.table import NaiveTableStyle
        editparam = NaiveTableStyle.from_entity(self)
        target = editparam.model_dump()
        registry_utils.register(
            self.__class__.__name__, key, target, extension=".json", disk=disk
        )

    def like(self, key):
        from fairypptx.editjson.table import NaiveTableStyle
        if isinstance(key, str):
            target = registry_utils.fetch(self.__class__.__name__, key)
            editparam = NaiveTableStyle.model_validate(target)
            editparam.apply(self)
            return self
        raise TypeError(f"Currently, type {type(style)} is not accepted.")




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
        writer = TableApiWriter(table.api)
        for r in range(n_row):
            for c in range(n_col):
                writer.write((r, c), arr[r, c])
        return table

    @staticmethod
    def empty(n_row: int = 3, n_col: int = 2) -> Table:
        shapes_api = resolve_shapes()
        shape_api = shapes_api.AddTable(NumRows=n_row, NumColumns=n_col)
        return Table(shape_api.Table)
        
        

if __name__ == "__main__":
    table = Table()
    array = np.zeros(shape=(4, 2)).astype(object)
    table[0:1, :2] = "TARET"
    array = np.array([[1, 2], [3, 4]])
    print(slice(0, 2))
    array[0:2, 1] = [0.5, 0.6]
    print(array)
