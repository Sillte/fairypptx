"""table class 

* `DFTable` enables you to get `pandas.DataFrame` from `Table` Object.
* 

Desire:
----------
* Currenly, `Table` / `DFTable` cannot be read from the existing table.
(2021-03-28) Well, then I'd like to attack this problem.

"""

from typing import Any, Self
from collections.abc import Sequence
import numpy as np
import pandas as pd
from fairypptx import constants
from fairypptx.constants import msoTrue, msoFalse

from fairypptx.color import Color
from fairypptx.core.application import Application
from fairypptx.core.resolvers import resolve_table, resolve_shapes
from fairypptx.core.types import COMObject
from fairypptx.slide import Slide
from fairypptx.shapes import Shapes
from fairypptx.object_utils import is_object, upstream, stored
from fairypptx.object_utils import registry_utils

from fairypptx._table import Cell, Row, Rows, Column, Columns
from fairypptx._table.stylist import TableStylist

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
        writer = TableWriter(self)
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
        data = [row.tolist() for row in self.rows]
        return data

    def to_numpy(self):
        """Convert to `numpy`.
        """
        data = self.tolist()
        print("data", data)
        return np.array(data)

    @property
    def values(self):
        return self.to_numpy()  

    def register(self, key, disk=True):
        stylist = TableStylist(self)
        registry_utils.register(
            self.__class__.__name__, key, stylist, extension=".pkl", disk=disk
        )

    def like(self, key):
        if isinstance(key, str):
            stylist = registry_utils.fetch(self.__class__.__name__, key)
            stylist(self)
            return self
        raise TypeError(f"Currently, type {type(style)} is not accepted.")


class TableWriter:
    """Encapsulates all writing logic for Table.
    """

    def __init__(self, table: Table):
        self.table = table

    # ---- Public API -------------------------------------------------

    def write_cell(self, r: int, c: int, value: Any) -> None:
        cell = self.table.api.Cell(r + 1, c + 1)
        cell.Shape.TextFrame.TextRange.Text = str(value)

    def write(self, key: tuple[int | list | slice, int | list | slice], value: np.ndarray | Any) -> None:
        """Implements the same logic as the current Table.__setitem__."""
        i_row, i_col = key

        # Case 1: scalar index
        if isinstance(i_row, int) and isinstance(i_col, int):
            r = self._normalize(i_row, axis=0)
            c = self._normalize(i_col, axis=1)
            self.write_cell(r, c, value)
            return

        # Case 2: broadcast cases
        r_idx = self._to_indices(i_row, axis=0)
        c_idx = self._to_indices(i_col, axis=1)

        value = np.array(value)

        # Broadcasting
        if value.ndim == 0:
            value = np.broadcast_to(value, (len(r_idx), len(c_idx)))
        elif value.ndim == 1:
            if len(r_idx) == 1:
                value = value[None, :]
            else:
                value = value[:, None]

        assert value.shape == (len(r_idx), len(c_idx))

        for i, rr in enumerate(r_idx):
            for j, cc in enumerate(c_idx):
                self.write_cell(rr, cc, value[i, j])

    # ---- Helpers -----------------------------------------------------

    def _normalize(self, idx, axis):
        size = self.table.size[axis]
        return idx + size if idx < 0 else idx

    def _is_indices(self, arg):
        return isinstance(arg, (list, tuple, slice))

    def _to_indices(self, arg, axis):
        if isinstance(arg, int):
            return [self._normalize(arg, axis)]

        if isinstance(arg, (list, tuple)):
            return [self._normalize(x, axis) for x in arg]

        if isinstance(arg, slice):
            start = arg.start or 0
            stop = arg.stop or self.table.size[axis]
            step = arg.step or 1
            return list(range(start, stop, step))

        raise ValueError(f"Invalid index specifier: {arg}")

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
    def from_numpy(arr: np.ndarray):
        arr = np.atleast_2d(arr)
        n_row, n_col = arr.shape

        shapes_api = resolve_shapes()
        shape_api = shapes_api.AddTable(NumRows=n_row, NumColumns=n_col)
        table = Table(shape_api.Table)

        writer = TableWriter(table)

        for r in range(n_row):
            for c in range(n_col):
                writer.write((r, c), arr[r, c])
        return table

    @staticmethod
    def empty(n_row: int = 3, n_col: int = 2):
        shapes_api = resolve_shapes()
        shape_api = shapes_api.AddTable(NumRows=n_row, NumColumns=n_col)
        return Table(shape_api.Table)
        
        


class DFTable:
    """`pandas.DataFrame` Table.
    This class is intended to handle `pandas.DataFrame`.  
    """
    def __init__(self,
                 arg=None,
                 *, index_nlevels=None,
                 columns_nlevels=None):

        self.app = Application()
        self._api = self._fetch_api(arg)
        self.index_nlevels = index_nlevels
        self.columns_nlevels = columns_nlevels

    def _fetch_api(self, arg):
        from fairypptx import Shape
        if is_object(arg, "Shape"):
            assert arg.Type == constants.msoTable
            return arg.Table
        elif isinstance(arg, Shape):
            assert arg.api.Type == constants.msoTable
            return arg.api.Table
        elif is_object(arg, "Table"):
            return arg
        elif isinstance(arg, DFTable):
            return arg.api
        elif arg is None:
            return self._fetch_api(Shape())

        raise ValueError(f"Cannot interpret `arg`; {arg}.")

    @property
    def api(self):
        return self._api

    @classmethod
    def make(self, df, index=True):
        assert isinstance(df, pd.DataFrame)

        if index is True:
            columns = df.columns.values
            index = df.index.values

            index_nlevels = df.index.nlevels
            column_nlevels = df.columns.nlevels
            n_row, n_column = df.shape

            table = Table.make(size=(column_nlevels + n_row, index_nlevels + n_column))

            # columns.values
            for i_level in range(column_nlevels):
                for index, value in enumerate(df.columns.get_level_values(i_level)):
                    table[i_level, index_nlevels + index] = value

            # index.values
            for i_level in range(index_nlevels):
                for index, value in enumerate(df.index.get_level_values(i_level)):
                    table[column_nlevels + index, i_level] = value

            # values
            for r_index in range(n_row):
                for c_index in range(n_column):
                    table[column_nlevels + r_index, index_nlevels + c_index] = df.iat[r_index, c_index]

            return DFTable(table.api, index_nlevels=index_nlevels, columns_nlevels=column_nlevels)
        else:
            index_nlevels = 0

            columns = df.columns.values
            column_nlevels = df.columns.nlevels
            n_row, n_column = df.shape

            table = Table.make(size=(column_nlevels + n_row,  n_column))

            # columns.values
            for i_level in range(column_nlevels):
                for index, value in enumerate(df.columns.get_level_values(i_level)):
                    table[i_level, index] = value

            # values
            for r_index in range(n_row):
                for c_index in range(n_column):
                    table[column_nlevels + r_index, c_index] = df.iat[r_index, c_index]

            return DFTable(table.api, index_nlevels=index_nlevels, columns_nlevels=column_nlevels)

    @property
    def size(self):
        # Naming is under consideration. `row` and `column` are more appropriate?
        return (len(self.api.Rows), len(self.api.Columns))

    @property
    def table(self):
        return Table(self.api)

    def tighten(self):
        self.table.tighten()

    def to_df(self, index_nlevels=None, columns_nlevels=None):
        """Return `pandas.DataFrame`.
        """
        if index_nlevels is None:
            index_nlevels = self.index_nlevels

        if columns_nlevels is None:
            columns_nlevels = self.columns_nlevels

        if index_nlevels is None or columns_nlevels is None:
            i_index_nlevels, i_columns_nlevels = self._infer_nlevels()
            if index_nlevels is None:
                index_nlevels = i_index_nlevels
            if columns_nlevels is None:
                columns_nlevels = i_columns_nlevels

        t_row, t_column = self.size
        array = self.table.to_numpy()

        # columns.values
        column_values = [ tuple([str(array[c_level,  c_index])
                          for c_level in range(columns_nlevels)])
                    for c_index in range(index_nlevels, t_column)]
        if 1 < columns_nlevels:
            columns = pd.MultiIndex.from_tuples(column_values)
        else:
            columns = [elem[0] for elem in column_values]
        # index.values
        index_values = [ tuple(str(array[i_index, i_level])
                         for i_level in range(index_nlevels))
                    for i_index in range(columns_nlevels, t_row)]

        if 1 < index_nlevels:
            index = pd.MultiIndex.from_tuples(index_values)
        elif 1 == index_nlevels:
            index = [elem[0] for elem in index_values]
        else:
            index = None

        # values
        n_row = t_row - columns_nlevels
        n_column = t_column - index_nlevels

        values = [[None] * n_column for _ in range(n_row)]
        for r_index in range(n_row):
            for c_index in range(n_column):
                values[r_index][c_index] = str(array[(r_index + columns_nlevels, c_index + index_nlevels)])

        df = pd.DataFrame(np.array(values), index=index, columns=columns)

        # Type inference via text and conversion.
        types = []
        for c_index in range(n_column):
            inferred = set((_TypeGuess.guess(values[r_index][c_index]) for r_index in range(n_row)))
            types.append(_TypeGuess.min(inferred))
        for c_index, t in enumerate(types): 
           df.iloc[:, c_index] = df.iloc[:, c_index].astype(t)
        return df


    @property
    def df(self):
        """Return 
        """

        """
        Note
        ------
         `shape.text` is not `str`, but `Text(UserString)`. 
        """
        return self.to_df()

    def _infer_nlevels(self):
        """Returns index_nlevels and columns_nlevels based on the contents of display. 

        Note: (2021-03-28) I feel there is much room for improvement.
        """
        def _is_content(arg):
            if str(arg).strip():
                return True
            return False
        first_row = self.table.rows[0].tolist()
        first_column = self.table.columns[0].tolist()

        for index, value in enumerate(first_row):
            if _is_content(value): 
                index_nlevels = index
                break
        else:
            index_nlevels = 0

        for index, value in enumerate(first_column[1:]):
            if _is_content(value): 
                columns_nlevels = index + 1
                break
        else:
            columns_nlevels = 1

        return index_nlevels, columns_nlevels

    def _yield_nlevels(self, index_nlevels=None, columns_nlevels=None):
        """Solves `index_nlevels` and `columns_nlevels`.

        1. If `index_nlevels` or `columns_nlevels` are clarified,  
        then they are used.
        2. If not, they are inferred via `self._infer_nlevels`. 
        """
        if self.index_nlevels is None or self.columns_nlevels is None:
            i_index_nlevels, i_columns_nlevels  = self._infer_nlevels()
            if self.index_nlevels is not None:
                index_nlevels = self.index_nlevels
            else:
                index_nlevels = i_index_nlevels
            if self.columns_nlevels is not None:
                columns_nlevels = self.columns_nlevels
            else:
                columns_nlevels = i_columns_nlevels
        else:
            index_nlevels = self.index_nlevels
            columns_nlevels = self.columns_nlevels
        return index_nlevels, columns_nlevels

    @property
    def index(self):
        return self.df.index

    @index.setter
    def index(self, values):
        index_nlevels, columns_nlevels = self._yield_nlevels()
        table = self.table 
        t_row = len(self.api.Rows)

        length = t_row - columns_nlevels

        if index_nlevels == 0:
            raise ValueError("This DFTable's index is empty.")
        values  = np.array(values)
        if values.ndim == 0:
            raise ValueError("Invalid")
        if values.ndim == 1:
            values = values[..., None]

        if values.shape[-1] != index_nlevels:
            raise ValueError("The level of index is different.", f"Given:{values.shape[-1]}, Table:{index_nlevels}")

        if values.shape[0] != length:
            raise ValueError("The length of index is different.", f"Given:{values.shape[0]}, Table:{length}")

        for r_index in range(length):
            for c_index in range(index_nlevels):
                table[columns_nlevels + r_index, c_index] = values[r_index, c_index]

    @property
    def columns(self):
        return self.df.columns

    @columns.setter
    def columns(self, values):
        index_nlevels, columns_nlevels = self._yield_nlevels()
        table = self.table 
        t_columns = len(self.api.Columns)

        length = t_columns - index_nlevels

        values  = np.array(values)
        if values.ndim == 0:
            raise ValueError("The dim of give values is 0.")
        if values.ndim == 1:
            values = values[..., None]
        if values.shape[-1] != columns_nlevels:
            raise ValueError("The level of columns is different.", f"Given:{values.shape[-1]}, Table:{columns_nlevels}")

        if values.shape[0] != length:
            raise ValueError("The length of columns is different.", f"Given:{values.shape[0]}, Table:{length}")


        for r_index in range(columns_nlevels):
            for c_index in range(length):
                table[r_index, index_nlevels + c_index] = values[c_index, r_index]

    def tolist(self):
        return self.df.tolist()

    def to_numpy(self):
        """Convert to `numpy`.
        """
        return self.df.to_numpy()

    @property
    def at(self):
        return _AtIndexer(self)


    def __getitem__(self, key):
        return self.df[key]

    def __setitem__(self, key, value):
        index_nlevels, columns_nlevels = self._yield_nlevels()
        if isinstance(key, (str, int, tuple)):
            ci = list(self.columns).index(key)
            self.table[columns_nlevels]



class _TypeGuess:

    # Order: from the highest (the best specific object) to the lowest (the most general object).  
    type_infos = [(int, int), (float, float), (str, str)]
    type_to_priority = {elem[0]: -p for p, elem in enumerate(type_infos)}

    @classmethod
    def guess(cls, arg):
        for type_info in cls.type_infos:
            type, call = type_info 
            try:
                call(arg)
            except ValueError:
                pass
            else:
                return type

        raise ValueError(f"Cannot guess the type of `arg`.")

    @classmethod
    def min(cls, types):
        """ Guess the most safe type over `types`.
        """
        return min(types, key=lambda t: cls.type_to_priority[t], default=str) 

class _AtIndexer:
    def __init__(self, df_table):
        self.df_table = df_table

    def __setitem__(self, key, value):
        ii, cc = self._to_indices(key)
        self.df_table.table[ii, cc] = value

    def __getitem__(self, key):
        ii, cc = self._to_indices(key)
        result = self.df_table.table[ii, cc].text
        return _TypeGuess.guess(result)(result)

    def _to_indices(self, key):
        index_nlevels, columns_nlevels = self.df_table._yield_nlevels()

        columns = self.df_table.columns
        i_key, c_key = key
        if index_nlevels != 0:
            i = list(self.df_table.index).index(i_key)
        else:
            i = i_key
        c = list(self.df_table.columns).index(c_key)
        return index_nlevels + i, columns_nlevels + c


if __name__ == "__main__":
    table = Table()
    array = np.zeros(shape=(4, 2)).astype(object)
    table[0:1, :2] = "TARET"
    array = np.array([[1, 2], [3, 4]])
    print(slice(0, 2))
    array[0:2, 1] = [0.5, 0.6]
    print(array)
    df_table = DFTable()
    print(df_table.df)
    exit(0)

    import numpy as np
    data = np.random.normal(size=(3, 2))
    df = pd.DataFrame(data, columns=["A", "B"])
    df.__setitem__
    df = df.round(2)
    df.at
    df_table = DFTable.make(df)
    print(time.time() - s)
    df_table.index = ["One", "Two", "Three"]
    df_table.columns = ["AA", "BB"]
    df_table.at["One", "BB"] = 12.5
    df_table.tighten()

    #table = Table(Shape())
    #Table.make(
    #values = table.tolist()
    #print(values)
    ##print(table.to_numpy())
    df = DFTable().df
    df.iloc
    print(df)
    print(df.columns)
    print(df.index)
    exit(0)

    df = pd.DataFrame(np.arange(12).reshape(3, 4))
    df.index = pd.MultiIndex.from_tuples([("ア", "A"), ("ア", "B"), ("ア", "C")])
    df.columns = ["W", "X", "Y", "Z"]
    table = DFTable.make(df)
    print(table.df)
    array = np.random.uniform(size=(2, 3))
    print(table.df)
    
