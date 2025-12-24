from fairypptx import registry_utils
from fairypptx import constants
from fairypptx.core.utils import get_discriminator_mapping
from fairypptx.states.context import Context
from fairypptx.constants import msoTrue, msoFalse
from fairypptx.shape import Shape, TableShape, GroupShape
from fairypptx.shape_range import ShapeRange
from fairypptx.box import Box 
from pydantic import Field
from typing import Annotated
from fairypptx.states.models import FrozenBaseStateModel, BaseStateModel
from fairypptx.states.models import FrozenBaseStateModel, BaseStateModel


class FrozenBaseShapeStateModel(FrozenBaseStateModel):
    box: Annotated[Box, Field(description="Represents the position of the shape")]  # (Note that this is jsonable).
    zorder: Annotated[int, Field(description="The value of Zorder")]

class BaseShapeStateModel(BaseStateModel):
    box: Annotated[Box, Field(description="Represents the position of the shape")]  # (Note that this is jsonable).
    zorder: Annotated[int, Field(description="The value of Zorder")]
