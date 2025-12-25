from fairypptx import Shape
from fairypptx.enums import MsoShapeType
from fairypptx.shape import Shape
from fairypptx.states.context import Context
from fairypptx.constants import msoShapeRectangle
from typing import TYPE_CHECKING


from pydantic import BaseModel


from typing import Self, get_args

if TYPE_CHECKING:
    from fairypptx.states.shape.elements import AutoShapeStateModel


class UnknownInvalidValueModel(BaseModel, frozen=True):
    @classmethod
    def from_shape(cls, shape: Shape) -> Self:
        print(f"UnknownInvalidValueModel @ from_shape, {shape=}")
        return cls()

    def create_shape(self, state_model: "AutoShapeStateModel", context: Context) -> Shape:
        shapes = context.slide.shapes
        shape = shapes.add(msoShapeRectangle)
        shape.text = "UnknownInvalid"
        shape.box = state_model.box
        return shape


class InvalidLineValueModel(BaseModel, frozen=True):
    begin_x: float
    begin_y: float
    end_x: float
    end_y: float

    @classmethod
    def predicator(cls, shape: Shape) -> bool:
        shape_api = shape.api
        if shape_api.Connector and shape_api.ConnectorFormat.Type == 1:
            return True
        return False

    @classmethod
    def from_shape(cls, shape: Shape) -> Self:
        l, t, w, h = shape.api.Left, shape.api.Top, shape.api.Width, shape.api.Height

        hf = shape.api.HorizontalFlip  # msoTrue = -1, msoFalse = 0
        vf = shape.api.VerticalFlip

        # X座標の決定
        if hf:
            begin_x, end_x = l + w, l
        else:
            begin_x, end_x = l, l + w
        # Y座標の決定
        if vf: # 下から上の場合
            begin_y, end_y = t + h, t

        else:  # 上から下の場合
            begin_y, end_y = t, t + h
        return cls(begin_x=begin_x, begin_y=begin_y, end_x=end_x, end_y=end_y)

    def create_shape(self, state_model: "AutoShapeStateModel", context: Context) -> Shape:
        shapes = context.slide.shapes
        shape_api = shapes.api.AddLine(self.begin_x, self.begin_y, self.end_x, self.end_y)
        shape = Shape(shape_api)
        shape.box = state_model.box
        state_model.line.apply(shape.line)
        return shape


class InvalidTextBoxValueModel(BaseModel, frozen=True):
    """TextBox, however, 
    """
    @classmethod
    def predicator(cls, shape: Shape) -> bool:
        shape_api = shape.api
        if shape_api.AutoShapeType == MsoShapeType.NotPrimitive and shape.text_frame.api.HasText:
            return True
        return False

    @classmethod
    def from_shape(cls, _: Shape) -> Self:
        return cls()

    def create_shape(self, state_model: "AutoShapeStateModel", context: Context) -> Shape:
        shapes = context.slide.shapes
        box = state_model.box
        shape_api = shapes.api.AddShape(Type=1, Left=box.left, Top=box.top, Width=box.width, Height=box.height)
        shape = Shape(shape_api)
        shape.box = state_model.box
        state_model.line.apply(shape.line)
        if state_model.fill.valid:
            state_model.fill.apply(shape.fill)
        state_model.text_frame.apply(shape.text_frame)
        return shape


InvalidAutoShapeValueModel = InvalidLineValueModel | InvalidTextBoxValueModel | UnknownInvalidValueModel


def get_invalid_value_model(shape: Shape) -> InvalidAutoShapeValueModel:
    candidates = get_args(InvalidAutoShapeValueModel)
    for candidate in candidates:
        if candidate is UnknownInvalidValueModel:
            continue
        if candidate.predicator(shape):
            return candidate.from_shape(shape)
    return UnknownInvalidValueModel()