from fairypptx.core.models import BaseApiModel
from fairypptx.core.types import COMObject
from fairypptx.apis.text_frame.api_model import TextFrameApiModel

from collections.abc import Mapping, Sequence
from typing import Any, ClassVar, Mapping, Self, Sequence

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

    @classmethod
    def from_api(cls, api: COMObject) -> Self:
        rows = RowsApiModel.from_api(api.Rows)
        return cls(rows=rows)

    def apply_api(self, api: COMObject) -> None:
        self.rows.apply_api(api.Rows)
