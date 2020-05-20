""" Experiment as for various things.
"""
from fairypptx import Shape, Slide
from fairypptx import Shape, Application
from fairypptx.object_utils import is_object
from fairypptx import TextRange
from fairypptx.table import Table
from fairypptx._table import Row, Rows
from pprint import pprint
from fairypptx import constants
from fairypptx import Application

#shape = Shape()
table = Table()
for column in table.columns:
    print(column.is_empty())
    column.Width = 200
table.columns.tighten()
