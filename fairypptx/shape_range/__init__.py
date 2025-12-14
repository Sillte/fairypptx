from typing import Self, Sequence, Iterator
from collections.abc import Sequence as SeqABC

from fairypptx.core.types import COMObject
from fairypptx.core.application import Application
from fairypptx.box import Box
from fairypptx import constants
from fairypptx.shape import Shape, GroupShape
from fairypptx.object_utils import is_object
from fairypptx.core.resolvers import resolve_shape_range

from fairypptx._shape.location import ShapesAdjuster, ShapesAligner, ClusterAligner
from fairypptx.shape_range.types import AlignCMD, AlignParam
from fairypptx.shape_range.model import ShapeRange
from fairypptx.shape_range.editors.cluster import ShapeCluster


from fairypptx.shape_range.types import AlignCMD, AlignParam
