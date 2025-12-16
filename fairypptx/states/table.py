from fairypptx.states.models import BaseValueModel
from fairypptx.table import Table 
from pydantic import Field
from typing import Annotated, Self
from fairypptx.styles.table import NaiveTableStyle
from fairypptx.apis.table.api_model import TableApiModel


class TableValueModel(BaseValueModel):
    style: Annotated[NaiveTableStyle, Field(description="It represents the style of Table")]
    body: Annotated[TableApiModel, Field(description="It represents the contents of Table")]

    @classmethod
    def from_object(cls, object: Table) -> Self:
        return cls(style=NaiveTableStyle.from_entity(object), body=TableApiModel.from_api(object.api))


    def apply(self, object: Table):
        table = object
        self.body.apply_api(table.api)
        self.style.apply(table)

    @property
    def n_rows(self) -> int:  
        return len(self.body.rows)

    @property
    def n_columns(self) -> int:  
        return len(self.body.rows[0])

