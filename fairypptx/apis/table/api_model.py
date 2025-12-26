from fairypptx.core.models import BaseApiModel
from fairypptx.core.types import COMObject
from fairypptx.apis.text_frame.api_model import TextFrameApiModel
from fairypptx.box import Box

from collections import defaultdict   
from typing import Any, ClassVar, Mapping, Self, Sequence, Annotated
from pydantic import Field, BaseModel


class CellMergeValue(BaseModel, frozen=True):
    """This value class represents the merge operation of 
    """
    start_row: Annotated[int, Field(description="The start row index of merge, `0` based.")]
    start_column: Annotated[int, Field(description="The start column index of merge, `0` based.")]
    n_rows: Annotated[int, Field(description="The number of rows for merge.")]
    n_columns: Annotated[int, Field(description="The number of columns for merge.")]

    @classmethod
    def from_positions(cls, positions: Sequence[tuple[int, int]]) -> Self:
        rs, cs = tuple(zip(*positions))
        start_row = min(rs)
        n_rows = max(rs) - start_row + 1
        start_column = min(cs)
        n_columns = max(cs) - start_column + 1
        assert len(positions) == n_rows * n_columns, "Rectangle positions."
        return cls(start_row=start_row, start_column=start_column, n_rows=n_rows, n_columns=n_columns)

    def apply_table_api(self, table_api: COMObject) -> None:
        start_cell = table_api.Cell(self.start_row + 1, self.start_column + 1)
        end_cell = table_api.Cell(self.start_row + self.n_rows, self.start_column + self.n_columns)
        start_cell.Merge(end_cell)

    @classmethod
    def merge(cls, table_api: COMObject,start_row: int, start_column: int, n_rows: int, n_columns: int):
        start_cell = table_api.Cell(start_row + 1, start_column + 1)
        end_cell = table_api.Cell(start_row + n_rows, start_column + n_columns)
        start_cell.Merge(end_cell)
        return 


class CellMergeValues(BaseModel, frozen=True):
    items: Annotated[Sequence[CellMergeValue], Field(description="Sequence of `CellMerge`.")]

    @classmethod
    def from_table_api(cls, table_api: COMObject) -> Self:
        n_rows = len(table_api.Rows)
        n_columns = len(table_api.Columns)
        box_to_positions = defaultdict(list)
        for r in range(n_rows):
            for c in range(n_columns):
                box = Box.from_api(table_api.Cell(r + 1, c + 1).Shape)
                box_to_positions[box].append((r, c))
        items = []
        for _, value in box_to_positions.items():
            if len(value) == 1:
                continue
            items.append(CellMergeValue.from_positions(value))
        return cls(items=items)

    @classmethod
    def unmerge_all(cls, table_api: COMObject) -> None:
        """Unmerge all the cells.
        """
        n_rows = len(table_api.Rows)
        n_columns = len(table_api.Columns)
        
        # 1. 結合状態を物理座標(Box)でグループ化する
        box_to_positions = defaultdict(list)
        for r in range(n_rows):
            for c in range(n_columns):
                box = Box.from_api(table_api.Cell(r + 1, c + 1).Shape)
                box_to_positions[box].append((r, c))
                
        # 2. 結合されている（複数ポジションを持つ）Boxに対してのみSplitを実行
        for positions in box_to_positions.values():
            if len(positions) > 1:
                rs, cs = tuple(zip(*positions))
                start_r, start_c = min(rs), min(cs)
                count_r = max(rs) - start_r + 1
                count_c = max(cs) - start_c + 1
                try:
                    table_api.Cell(start_r + 1, start_c + 1).Split(count_r, count_c)
                except Exception as e:
                    print(f"Error at ({start_r}, {start_c}): {e} in Table.")


    def apply_table_api(self, table_api: COMObject):
        target = CellMergeValues.from_table_api(table_api)
        if target == self:
            return 
        items = sorted(self.items, key=lambda value: (value.start_row, value.start_column), reverse=True)
        for merge in items:
            merge.apply_table_api(table_api)
    

def normalize_index(api: COMObject, index: int | slice | Sequence[int]) -> int | Sequence[int]:
    def _normalize_int(index: int) -> int:
        return index % (api.Count)
    if isinstance(index, int):
        return _normalize_int(index)
    elif isinstance(index, slice):
        indices = list(range(*index.indices(api.Count)))
        return indices
    elif isinstance(index, Sequence):
        return [_normalize_int(elem) for elem in index]
    raise TypeError(f"Invalid Argument; `{index}`")

def insert_cell(api: COMObject, index: int) -> None:
    n_index = normalize_index(api, index)
    assert isinstance(n_index, int)
    api.Add(n_index + 1)
 

class CellApiModel(BaseApiModel):
    # Acutually, `Boarders` are required.
    text_frame: TextFrameApiModel
    @classmethod
    def from_api(cls, api: COMObject) -> Self:
        text_frame = TextFrameApiModel.from_api(api.Shape.TextFrame)
        return cls(text_frame=text_frame)

    def apply_api(self, api: COMObject) -> None:
        self.text_frame.apply_api(api.Shape.TextFrame)
        
        
def _apply_cells_to_api(row_or_column_api: COMObject, cells: Sequence[CellApiModel]):
    Cells = row_or_column_api.Cells
    desired = len(cells)
    actual = len(Cells)

    for _ in range(actual, desired):
        Cells.Add()
    for i in range(actual, desired, -1):
        Cells(i).Delete()

    for CellApi, cell in zip(Cells, cells):
        cell.apply_api(CellApi)
    return row_or_column_api


class RowApiModel(BaseApiModel):
    cells: Sequence[CellApiModel]
    height: float 
    @classmethod
    def from_api(cls, api: COMObject) -> Self:
        cells = [CellApiModel.from_api(elem) for elem in api.Cells]
        return cls(cells=cells, height=api.Height)

    def apply_api(self, api: COMObject):
        api.Height = self.height
        _apply_cells_to_api(api, self.cells)

    def __len__(self) -> int:
        return len(self.cells)


class ColumnApiModel(BaseApiModel):
    cells: Sequence[CellApiModel]
    width: float 
    @classmethod
    def from_api(cls, api: COMObject) -> Self:
        cells = [CellApiModel.from_api(elem) for elem in api.Cells]
        return cls(cells=cells, width=api.Width)

    def apply_api(self, api: COMObject):
        api.Width = self.width
        _apply_cells_to_api(api, self.cells)
        

    def __len__(self) -> int:
        return len(self.cells)

def _apply_collection_to_api(row_or_column_api, collection: Sequence[RowApiModel | ColumnApiModel]):
    api = row_or_column_api
    Item = api.Item
    actual = len(row_or_column_api)
    desired = len(collection)

    for _ in range(actual, desired):
        Item.Add()
    for i in range(actual, desired, -1):
        Item(i).Delete()

    for i, item in enumerate(collection, start=1):
        item.apply_api(Item(i))

    return row_or_column_api

class RowsApiModel(BaseApiModel):
    rows: Sequence[RowApiModel]
    @classmethod
    def from_api(cls, api: COMObject) -> Self:
        rows = [RowApiModel.from_api(elem) for elem in api]
        return cls(rows=rows)

    def apply_api(self, api: COMObject) -> None:
        _apply_collection_to_api(api, self.rows)

    def __len__(self) -> int:
        return len(self.rows)

    def __getitem__(self, index: int) -> RowApiModel:
        return self.rows[index]
    

class ColumnsApiModel(BaseApiModel):
    columns: Sequence[ColumnApiModel]
    @classmethod
    def from_api(cls, api: COMObject) -> Self:
        columns = [ColumnApiModel.from_api(elem) for elem in api]
        return cls(columns=columns)

    def apply_api(self, api: COMObject) -> None:
        _apply_collection_to_api(api, self.columns)

    def __len__(self) -> int:
        return len(self.columns)

    def __getitem__(self, index: int) -> ColumnApiModel:
        return self.columns[index]


class TableApiModel(BaseApiModel):
    rows: RowsApiModel
    merge_values: CellMergeValues

    @classmethod
    def from_api(cls, api: COMObject) -> Self:
        rows = RowsApiModel.from_api(api.Rows)
        merge_values = CellMergeValues.from_table_api(api)
        return cls(rows=rows, merge_values=merge_values)

    def apply_api(self, api: COMObject) -> None:
        self.merge_values.apply_table_api(api)
        self.rows.apply_api(api.Rows)
