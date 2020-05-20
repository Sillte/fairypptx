"""Table class 

Features
---------------------------
* 

"""

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
from fairypptx.inner import storage 
from fairypptx.shape import  Shapes, Shape
from fairypptx.object_utils import is_object, upstream, stored

from fairypptx._table import Cell, Row, Rows, Column, Columns

class Table:
    def __init__(self, arg=None, * ,app=None):
        if app is None:
            self.app = Application()
        else:
            self.app = app

        self._api = self._fetch_api(arg)

    def _fetch_api(self, arg):
        if is_object(arg, "Shape"):
            assert arg.Type == constants.msoTable
            return arg.Table
        elif isinstance(arg, Shape):
            assert arg.api.Type == constants.msoTable
            return arg.api.Table
        elif is_object(arg, "Table"):
            return arg
        elif isinstance(arg, Table):
            return arg.api
        elif arg is None:
            return self._fetch_api(Shape())

        raise ValueError(f"Cannot interpret `arg`; {arg}.")

    @property
    def shape(self):
        return Shape(self.api.Parent)

    @property
    def api(self):
        return self._api

    @property
    def size(self):
        # Naming is under consideration. `row` and `column` are more appropriate?
        return (len(self.api.Rows), len(self.api.Columns))

    @property
    def rows(self):
        return Rows(self.api.Rows)

    @property
    def columns(self):
        return Columns(self.api.Columns)

    def __setitem__(self, key, value):
        i_row, i_column = key
        cell_object = self.api.Cell(i_row + 1, i_column + 1)
        cell_object.Shape.TextFrame.TextRange.Text = str(value)

    def __getitem__(self, key):
        i_row, i_column = key
        cell_object = self.api.Cell(i_row + 1, i_column + 1)
        return Cell(cell_object)

    @classmethod
    def make(cls, arg=None, **kwargs):
        shapes = Shapes()

        if isinstance(arg, np.ndarray):
            assert arg.ndim <= 2
            arg = np.atleast_2d(arg)
            n_row, n_column = arg.shape
            shape_object = shapes.api.AddTable(NumRows=n_row, NumColumns=n_column)
            table = Table(shape_object.Table)
            
            for i_row in range(n_row):
                for i_column in range(n_column):
                    table[i_row, i_column] = str(arg[i_row, i_column])
            return table

        elif arg is None:
            n_row, n_column = kwargs.get("size", (1, 1))
            shape_object = shapes.api.AddTable(NumRows=n_row, NumColumns=n_column)
            table = Table(shape_object.Table)
            return table


        raise ValueError(f" `{arg}` is not interpretted.")

    @classmethod
    def empty(cls, shape, **kwargs):
        """From `np.empty`.
        """
        if isinstance(shape, int):
            row, col = shape, 1
        else:
            row, col = shape
        shape_object = Shapes().api.AddTable(NumRows=row, NumColumns=col)
        return Table(shape_object.Table)



class DFTable:
    """`pandas.DataFrame` Table.
    This class is intended to handle `pandas.DataFrame`.  
    """
    def __init__(self, arg=None, **kwargs):
        self.table, self.index_nlevels, self.column_nlevels = self._construct(arg, **kwargs)

    @property
    def df(self):
        """Return 
        """

        """
        Note
        ------
         `shape.text` is not `str`, but `Text(UserString)`. 
        """

        t_row, t_column = self.table.size

        # columns.values
        column_values = [ tuple([str(self.table[c_level,  c_index].text)
                          for c_level in range(self.column_nlevels)])
                    for c_index in range(self.index_nlevels, t_column)]
        if 1 < self.column_nlevels:
            columns = pd.MultiIndex.from_tuples(column_values)
        else:
            columns = [elem[0] for elem in column_values]
        # index.values
        index_values = [ tuple(str(self.table[i_index, i_level].text)
                         for i_level in range(self.index_nlevels))
                    for i_index in range(self.column_nlevels, t_row)]

        if 1 < self.index_nlevels:
            index = pd.MultiIndex.from_tuples(index_values)
        else:
            index = [elem[0] for elem in index_values]

        # values
        n_row = t_row - self.column_nlevels
        n_column = t_column - self.index_nlevels

        values = [[None] * n_column for _ in range(n_row)]
        for r_index in range(n_row):
            for c_index in range(n_column):
                values[r_index][c_index] = str(self.table[(r_index + self.column_nlevels, c_index + self.index_nlevels)].text)

        df = pd.DataFrame(np.array(values), index=index, columns=columns)

        # Type inference via text and conversion.
        types = []
        for c_index in range(n_column):
            inferred = set((_TypeGuess.guess(values[r_index][c_index]) for r_index in range(n_row)))
            types.append(_TypeGuess.min(inferred))
        for c_index, t in enumerate(types): 
            df = df.astype({df.columns.values[c_index]: t})
        return df


    def _construct(self, arg, **kwargs):
        assert isinstance(arg, pd.DataFrame)
        df = arg
        
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

        return table, index_nlevels, column_nlevels

        raise ValueError(f" `{arg}` is not interpretted.")



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


if __name__ == "__main__":
    df = pd.DataFrame(np.arange(12).reshape(3, 4))
    df.index = pd.MultiIndex.from_tuples([("ア", "A"), ("ア", "B"), ("ア", "C")])
    df.columns = ["W", "X", "Y", "Z"]
    table = DFTable(df)
    print(table.df)
    #array = np.random.uniform(size=(2, 3))
    #table = Table.make(array)
    #print(table[0, 0].text)
