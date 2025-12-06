from fairypptx.core.resolvers import resolve_table
from fairypptx.table import Table
from typing import Any 


import numpy as np
import pandas as pd

class ValueTypeGuess:
    _types = [int, float, str]

    @classmethod
    def guess_one(cls, val):
        for t in cls._types:
            try:
                t(val)
                return t
            except Exception:
                pass
        return str

    @classmethod
    def from_series(cls, series: pd.Series):
        guessed = {cls.guess_one(v) for v in series}
        # choose the "lowest" type (str > float > int)
        if str in guessed:
            return str
        if float in guessed:
            return float
        return int


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
    
class DFTableFrameReader:
    def __init__(self, table: Table):
        # table_api = COM Table object
        self.table = table

    def infer_nlevels(self) -> tuple[int, int]:
        """
        Inspect the first row and first column to infer header depths.
        Logic is based on the original DFTable._infer_nlevels.
        """
        array = self.table.to_numpy()

        def is_content(x: Any) -> bool:
            return bool(str(x).strip())

        index_nlevels = 0
        for i, v in enumerate(array[0, :]):
            if is_content(v):
                index_nlevels = i
                break
        # detect columns_nlevels (horizontal header depth)
        columns_nlevels = 1
        first_col = array[:, 0]
        for i, v in enumerate(first_col[1:], start=1):
            if is_content(v):
                columns_nlevels = i
                break
        return index_nlevels, columns_nlevels
    
    def to_dataframe(
        self,
        *,
        index_nlevels: int | None = None,
        columns_nlevels: int | None  = None,
    ) -> pd.DataFrame:
        """
           Convert PPT table into pandas.DataFrame.
        """
        # step 1: get raw text array
        array = self.table.to_numpy()
        n_rows, n_cols = array.shape

        # step 2: determine nlevels
        if index_nlevels is None or columns_nlevels is None:
            detected_index, detected_cols = self.infer_nlevels()
            index_nlevels = index_nlevels or detected_index
            columns_nlevels = columns_nlevels or detected_cols

        # step 3: reconstruct columns
        column_tuples = [
            tuple(
                str(array[row_level, col])
                for row_level in range(columns_nlevels)
            )
            for col in range(index_nlevels, n_cols)
        ]

        if columns_nlevels > 1:
            columns = pd.MultiIndex.from_tuples(column_tuples)
        else:
            columns = [t[0] for t in column_tuples]

        # step 4: reconstruct index
        index_tuples = [
            tuple(
                str(array[row, col_level])
                for col_level in range(index_nlevels)
            )
            for row in range(columns_nlevels, n_rows)
        ]

        if index_nlevels > 1:
            index = pd.MultiIndex.from_tuples(index_tuples)
        elif index_nlevels == 1:
            index = [t[0] for t in index_tuples]
        else:
            index = None

        # step 5: extract values (all str)
        body_rows = n_rows - columns_nlevels
        body_cols = n_cols - index_nlevels

        values = [
            [
                str(array[r + columns_nlevels, c + index_nlevels])
                for c in range(body_cols)
            ]
            for r in range(body_rows)
        ]

        df = pd.DataFrame(values, index=index, columns=columns)

        # step 6: infer value types for each column and cast
        for col_i in range(body_cols):
            col_series = df.iloc[:, col_i]
            inferred_type = ValueTypeGuess.from_series(col_series)
            df.iloc[:, col_i] = col_series.astype(inferred_type)

        return df
        


class DFTable:
    """`pandas.DataFrame` Table.
    This class is intended to handle `pandas.DataFrame`.  
    """
    def __init__(self,
                 arg=None,
                 *, index_nlevels: int | None =None,
                 columns_nlevels: int  | None =None):

        self._api = resolve_table(arg)
        self.index_nlevels = index_nlevels
        self.columns_nlevels = columns_nlevels

    @property
    def api(self):
        return self._api

    @classmethod
    def make(self, df: pd.DataFrame, with_index: bool = False):
        assert isinstance(df, pd.DataFrame)

        if with_index is True:
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

    def to_df(self, index_nlevels=None, columns_nlevels=None) -> pd.DataFrame:
        """Return `pandas.DataFrame`.
        """
        return DFTableFrameReader(self.table).to_dataframe(index_nlevels=index_nlevels, columns_nlevels=columns_nlevels)

    @property
    def df(self) -> pd.DataFrame:
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

    def _yield_nlevels(self, index_nlevels: int | None=None, columns_nlevels: int | None = None):
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
    def index(self) -> pd.Index:
        return self.df.index

    @index.setter
    def index(self, values):
        reader = DFTableFrameReader(self.table)
        index_nlevels, columns_nlevels = reader.infer_nlevels()
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
                self.table[columns_nlevels + r_index, c_index] = values[r_index, c_index]

    @property
    def columns(self) -> pd.Index:
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
                self.table[r_index, index_nlevels + c_index] = values[c_index, r_index]

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
            

if __name__ == "__main__":
    df_table = DFTable()
    print(df_table.df)
    exit(0)

    import numpy as np
    data = np.random.normal(size=(3, 2))
    df = pd.DataFrame(data, columns=["A", "B"])
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
    
