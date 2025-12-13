from fairypptx.core.models import ApiApplicator, COMObject
from fairypptx.apis.table.api_model import TableApiModel
from fairypptx.apis.table.api_model import RowsApiModel
from fairypptx.apis.table.api_model import RowApiModel
from fairypptx.apis.table.api_model import ColumnsApiModel
from fairypptx.apis.table.api_model import ColumnApiModel
from fairypptx.apis.table.api_model import CellApiModel

# No custom conversion function at the moment; use model-driven apply

RowApiApplicator = ApiApplicator(RowApiModel, None)
ColumnApiApplicator = ApiApplicator(ColumnApiModel, None)
RowsApiApplicator = ApiApplicator(RowsApiModel, None)
ColumnsApiApplicator = ApiApplicator(ColumnsApiModel, None)
TableApiApplicator = ApiApplicator(TableApiModel, None)
CellApiApplicator = ApiApplicator(CellApiModel, None)
