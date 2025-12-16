from pydantic import BaseModel 
from typing import Self, Any, Sequence, Mapping, ClassVar
from fairypptx import Shape
from fairypptx.core.utils import crude_api_read
from fairypptx.shape import Box
from typing import Protocol
from fairypptx.table import Table, Cell
from fairypptx.styles.text_range import NaiveTextRangeParagraphStyle
from fairypptx.core.utils import crude_api_write   

# * Generate the parameters for `ParamItself`.
# * Apply the generate params for Shape. 

class NaiveCellStyle(BaseModel):
    textrange_style: NaiveTextRangeParagraphStyle

    @classmethod
    def from_entity(cls, entity: Cell) -> Self:
        cell = entity
        textrange_style = NaiveTextRangeParagraphStyle.from_entity(cell.shape.textrange)
        return cls(textrange_style=textrange_style)

    def apply(self, entity: Cell):
        cell = entity
        self.textrange_style.apply(cell.shape.textrange)
        return cell

class NaiveTableStyle(BaseModel):
    """
    Naive table style.

    - Cell styles are sampled from existing cells.
    - If table size is larger than the stored mapping,
      styles are inherited from the nearest existing cell.
    """
    cell_style_mapping_pair: Sequence[tuple[tuple[int, int], NaiveCellStyle]]
    style_id: str | None = None
    api_data: Mapping[str, Any]

    _keys: ClassVar[Sequence[str]] = ["FirstCol", "FirstRow", "LastCol", "LastRow"]

    @classmethod
    def _to_mapping_pair(cls, cell_style_mapping: Mapping[tuple[int, int], NaiveCellStyle]) -> Sequence[tuple[tuple[int, int], NaiveCellStyle]]:
        return [(key, value) for key, value in cell_style_mapping.items()]

    @classmethod
    def _to_mapping(cls, cell_style_mapping_pair: Sequence[tuple[tuple[int, int], NaiveCellStyle]] ) -> Mapping[tuple[int, int], NaiveCellStyle]:
        return {elem[0]: elem[1] for elem in cell_style_mapping_pair}


    @classmethod
    def from_entity(cls, entity: Table) -> Self:
        table = entity
        data = crude_api_read(table.api, cls._keys)
        style_id = table.api.Style.Id
        cell_style_mapping = dict()
        for i, row in enumerate(table.rows):
            for j, cell in enumerate(row):
                cell_style_mapping[(i, j)] = NaiveCellStyle.from_entity(cell)
        pair = cls._to_mapping_pair(cell_style_mapping)
        return cls(cell_style_mapping_pair=pair, style_id=style_id, api_data=data)


    def apply(self, entity: Table) -> Table:
        table = entity
        if self.style_id:
            table.api.ApplyStyle(self.style_id)
        crude_api_write(table.api, self.api_data)

        cell_style_mapping = self._to_mapping(self.cell_style_mapping_pair)
        keys = cell_style_mapping.keys()
        max_i = max(key[0] for key in keys)
        max_j = max(key[1] for key in keys)
        for i, row in enumerate(table.rows):
            for j, cell in enumerate(row):
                if (i, j) in cell_style_mapping:
                    cell_style_mapping[i, j].apply(cell)
                else:
                    fallback_i = i if i <= max_i else max_i
                    fallback_j = j if j <= max_j else max_j
                    cell_style_mapping[fallback_i, fallback_j].apply(cell)
        return table


