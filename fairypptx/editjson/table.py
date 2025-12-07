from pydantic import BaseModel 
from typing import Self, Any, Sequence, Mapping, ClassVar
from fairypptx import Shape
from fairypptx.core.utils import crude_api_read
from fairypptx.shape import Box
from typing import Protocol
from fairypptx.table import Table
from fairypptx.editjson.text_range import NaiveTextRangeParagraphStyle
from fairypptx.core.utils import crude_api_write   

# * Generate the parameters for `ParamItself`.
# * Apply the generate params for Shape. 

class NaiveTableStyle(BaseModel):
    textrange_style: NaiveTextRangeParagraphStyle
    style_id: str | None = None
    api_data: Mapping[str, Any]

    _keys: ClassVar[Sequence[str]] = ["FirstCol", "FirstRow", "FirstCol", "LastCol"]

    @classmethod
    def from_entity(cls, entity: Table) -> Self:
        table = entity
        data = crude_api_read(table.api, cls._keys)
        style_id = table.api.Style.Id
        textrange_style = NaiveTextRangeParagraphStyle.from_entity(table.rows[0][0].shape.textrange)
        return cls(textrange_style=textrange_style, style_id=style_id, api_data=data)


    def apply(self, entity: Table) -> Table:
        table = entity
        if self.style_id:
            table.api.ApplyStyle(self.style_id)
        crude_api_write(table.api, self.api_data)
        for row in table.rows:
            for cell in row:
                self.textrange_style.apply(cell.shape.textrange)
        return table


