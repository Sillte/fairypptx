from fairypptx import constants 
from fairypptx._text.textrange_stylist import ParagraphTextRangeStylist
from fairypptx._shape.stylist import ShapeStylist
from fairypptx.object_utils import ObjectDictMixin
from fairypptx import constants


class _BasicTableFormat(ObjectDictMixin):
    # The Object Class name, which is the staring point.
    name = "Table"
    data = dict()
    attributes = ["FirstCol", "FirstRow", ]
    data["FirstRow"] = constants.msoTrue
    data["FirstCol"] = constants.msoFalse
    data["LastRow"] = constants.msoFalse
    data["LastCol"] = constants.msoFalse
    data["Style.ID"] = "{22838BEF-8BB2-4498-84A7-C5851F593DF1}"  
    # Special handling is necessary.
    readonly = ["Style.ID"]  

    def apply(self, api_object):
        super().apply(api_object)
        api_object.ApplyStyle(self.data["Style.ID"])


class TableStylist:
    def __init__(self, table):
        self.table_format = _BasicTableFormat(table.api)
        self.shape_stylelist = ShapeStylist(table.rows[0][0].shape)

    def __call__(self, table):
        self.table_format.apply(table.api)

        for row in table.rows:
            for cell in row:
                self.shape_stylelist(cell.shape)
        return table


if __name__ == "__main__":
    from fairypptx import Table
    TableStylist(Table())
    TableStylist(Table())(Table())
