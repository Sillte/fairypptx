"""Table API layer exports.

Exports:
  - TableApiModel, TableApplicator: Top-level table model and applicator
  - RowsApiModel, RowsApiApplicator: Collection of rows
  - ColumnsApiModel, ColumnsApiApplicator: Collection of columns
  - RowApiModel, RowApiApplicator: Single row with cells
  - ColumnApiModel, ColumnApiApplicator: Single column with cells
  - CellApiModel: Single cell (text frame + formatting)
"""

from fairypptx.apis.table.api_model import (
    TableApiModel,
    RowsApiModel,
    RowApiModel,
    ColumnsApiModel,
    ColumnApiModel,
    CellApiModel,
)
from fairypptx.apis.table.applicator import (
    TableApiApplicator,
    RowsApiApplicator,
    RowApiApplicator,
    ColumnsApiApplicator,
    ColumnApiApplicator,
    CellApiApplicator,
)

__all__ = [
    "TableApiModel",
    "TableApiApplicator",
    "RowsApiModel",
    "RowsApiApplicator",
    "RowApiModel",
    "RowApiApplicator",
    "ColumnsApiModel",
    "ColumnsApiApplicator",
    "ColumnApiModel",
    "ColumnApiApplicator",
    "CellApiModel",
    "CellApiApplicator",
]
