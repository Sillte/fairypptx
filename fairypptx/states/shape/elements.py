import io
import base64
from pydantic import BaseModel 
from PIL import Image
from fairypptx import Shape,  constants, registry_utils
from fairypptx.shape import TableShape
from fairypptx.box import Box
from fairypptx.constants import msoFalse, msoTrue
from fairypptx.enums import MsoShapeType
from fairypptx.shape import Shape
from fairypptx.shapes import Shapes
from fairypptx.states.context import Context
from fairypptx.states.shape.base import BaseShapeStateModel, FrozenBaseShapeStateModel
from fairypptx.states.table import TableValueModel
from fairypptx.states.text_frame import TextFrameValueModel
from fairypptx.styles.fill_format import NaiveFillFormatStyle
from fairypptx.styles.line_format import NaiveLineFormatStyle

from pydantic import Base64Bytes, Field

from typing import Annotated, Literal, Self, Any, Mapping

from fairypptx.table import Table, cast


class AutoShapeStateModel(FrozenBaseShapeStateModel):
    type: Annotated[Literal[MsoShapeType.AutoShape], Field(description="Type of Shape")] = MsoShapeType.AutoShape
    style_index: Annotated[int | None, Field(description="Style of Shape")]
    auto_shape_type: Annotated[int, Field(description="Represents MSOAutoShapeType.")]
    line: Annotated[NaiveLineFormatStyle, Field(description="Represents the format of `Line` around the Shape.")]
    fill: Annotated[NaiveFillFormatStyle, Field(description="Represents the format of `Fill` of the Shape.")]
    text_frame: Annotated[TextFrameValueModel, Field(description="Represents the texts of the Shape.")]

    @classmethod
    def from_entity(cls, entity: Shape) -> Self:
        shape = entity
        return cls(box=shape.box,
                   id=shape.id,
                   style_index=shape.style_index,
                   auto_shape_type=shape.api.AutoShapeType,
                   line=NaiveLineFormatStyle.from_entity(shape.line),
                   fill=NaiveFillFormatStyle.from_entity(shape.fill),
                   text_frame=TextFrameValueModel.from_object(shape.text_frame),
                   zorder=shape.api.ZOrderPosition,
                   )

    def create_entity(self, context: Context) -> Shape:
        shapes = context.shapes
        shape = shapes.add(auto_shape_type=self.auto_shape_type)
        self.apply(shape)
        return shape

    def apply(self, entity: Shape) -> Shape:
        shape = entity
        shape.box = self.box

        if self.auto_shape_type not in {constants.msoShapeNotPrimitive}:
            shape.api.AutoShapeType = self.auto_shape_type

        self.text_frame.apply(shape.text_frame)

        shape.style_index = self.style_index
        self.line.apply(shape.line)
        if self.fill.valid:
            self.fill.apply(shape.fill)
        elif self._should_clear_fill():
            shape.fill = None
        return shape

    def _should_clear_fill(self) -> bool:
        return (not self.fill.valid) and (not self.style_index)


class TableShapeStateModel(FrozenBaseShapeStateModel):
    type: Annotated[Literal[MsoShapeType.Table], Field(description="Type of Shape")] = MsoShapeType.Table
    table: Annotated[TableValueModel, Field(description="Table of the Shape")]
    @classmethod
    def from_entity(cls, entity: Shape) -> Self:
        shape = cast(TableShape, entity)
        return cls(box=shape.box,
                   id=shape.id,
                   table=TableValueModel.from_object(shape.table),
                   zorder=shape.api.ZOrderPosition,
                   )

    def create_entity(self, context: Context) -> Shape:
        shapes = context.shapes
        n_rows = self.table.n_rows
        n_columns = self.table.n_columns
        shape_api = shapes.api.AddTable(NumRows=n_rows, NumColumns=n_columns)
        table = Table(shape_api.Table)
        shape = table.shape
        self.apply(shape)
        return table.shape

    def apply(self, entity: Shape) -> Shape:
        shape = cast(TableShape, entity)
        shape.box = self.box
        self.table.apply(shape.table)
        return shape


class PictureShapeStateModel(BaseShapeStateModel):
    type: Annotated[Literal[MsoShapeType.Picture], Field(description="Type of Shape")] = MsoShapeType.Picture
    box: Annotated[Box, Field(description="Represents the position of the shape")]  # (Note that this is jsonable).
    image: Annotated[Base64Bytes, Field(description="Image of the shape.")]
    zorder: Annotated[int, Field(description="The value of Zorder")]

    @classmethod
    def from_entity(cls, entity: Shape) -> Self:
        shape = entity
        image=shape.to_image()
        buffer = io.BytesIO()
        image.save(buffer, format="PNG")
        image_bytes = buffer.getvalue()
        return cls(box=shape.box,
                   id=shape.id,
                   image=base64.b64encode(image_bytes),
                   zorder=shape.api.ZOrderPosition,
                   )

    def create_entity(self, context: Context) -> Shape:
        shapes_api = context.shapes.api
        img_data = self.image
        image = Image.open(io.BytesIO(img_data))

        with registry_utils.yield_temporary_dump(image, suffix=".png") as path:
            shape = Shape(shapes_api.AddPicture(str(path), msoFalse, msoTrue, Left=self.box.left, Width=self.box.width, Top=self.box.top, Height=self.box.height))
        return shape

    def apply(self, entity: Shape) -> Shape:
        orig_shape = entity
        context = Context(slide=orig_shape.slide)
        shape = self.create_entity(context)
        shape.box = self.box
        self.id = shape.id # NOTE: Since the new object is created, this is ineviable.
        orig_shape.api.Delete()
        return shape


class TextBoxShapeStateModel(FrozenBaseShapeStateModel):
    type: Annotated[Literal[MsoShapeType.TextBox], Field(description="Type of Shape")] = MsoShapeType.TextBox
    line: Annotated[NaiveLineFormatStyle, Field(description="Represents the format of `Line` around the Shape.")]
    fill: Annotated[NaiveFillFormatStyle, Field(description="Represents the format of `Fill` of the Shape.")]
    text_frame: Annotated[TextFrameValueModel, Field(description="Represents the texts of the Shape.")]

    @classmethod
    def from_entity(cls, entity: Shape) -> Self:
        shape = entity
        return cls(box=shape.box,
                   id=shape.id,
                   line=NaiveLineFormatStyle.from_entity(shape.line),
                   fill=NaiveFillFormatStyle.from_entity(shape.fill),
                   text_frame=TextFrameValueModel.from_object(shape.text_frame),
                   zorder=shape.api.ZOrderPosition,
                   )

    def create_entity(self, context: Context) -> Shape:
        shapes_api = context.shapes.api
        shape_api = shapes_api.AddTextbox(constants.msoTextOrientationHorizontal, Left=self.box.left, Top=self.box.top, Width=self.box.width, Height=self.box.height)
        shape = Shape(shape_api)
        self.apply(shape)
        return shape

    def apply(self, entity: Shape) -> Shape:
        shape = entity
        shape.box = self.box
        self.line.apply(shape.line)
        self.fill.apply(shape.fill)
        self.text_frame.apply(shape.text_frame)
        return shape

class ConnectPair(BaseModel, frozen=True):
    id: Annotated[int, Field(description="Shape Id")] 
    site: Annotated[int, Field(description="Site number")] 

class ConnectorFormatValue(BaseModel, frozen=True):
    begin: Annotated[ConnectPair | None, Field(description="ShapeId of Start")] 
    end: Annotated[ConnectPair | None, Field(description="Number of Site")] 

    @classmethod
    def from_shape_api(cls, shape_api: Any) -> Self: 
        assert shape_api.Connector
        c_format = shape_api.ConnectorFormat
        if c_format.BeginConnected:
            begin_id = c_format.BeginConnectedShape.Id
            begin_site = c_format.BeginConnectionSite
            begin = ConnectPair(id=begin_id, site=begin_site)
        else:
            begin = None

        if c_format.EndConnected:
            end_id = c_format.EndConnectedShape.Id
            end_site = c_format.EndConnectionSite
            end = ConnectPair(id=end_id, site=end_site)
        else:
            end = None
        return cls(begin=begin, end=end)

    def apply(self, line_shape: Shape, id_mapping:Mapping[int, int], shapes: Shapes) -> None: 
        """Set the connector to `line_shape`, referring to `id_mapping` and `shapes`.
        The stored `id` in `ConnectPair` is resolved with `id_mapping` for the ids of `Shapes`.
        """
        id_to_shape = {shape.id: shape for shape in shapes} 
        c_format = line_shape.api.ConnectorFormat
        try:
            if self.begin:
                shape_id = id_mapping[self.begin.id]
                shape = id_to_shape[shape_id]
                c_format.BeginConnect(shape.api, self.begin.site)
            if self.end:
                shape_id = id_mapping[self.end.id]
                shape = id_to_shape[shape_id]
                c_format.EndConnect(shape.api, self.end.site)
        except KeyError as e:
            print(e, "In ConnectorFormatValue, `connection` is tried, but Id is not resolved correctly.")



class LineShapeStateModel(FrozenBaseShapeStateModel):
    type: Annotated[
        Literal[MsoShapeType.Line],
        Field(description="Type of Shape")
    ] = MsoShapeType.Line

    line: Annotated[
        NaiveLineFormatStyle,
        Field(description="Represents the format of Line")
    ]

    connector: Annotated[ConnectorFormatValue | None, Field(description="Connector Format")] = None

    @classmethod
    def from_entity(cls, entity: Shape) -> Self:
        shape = entity
        if shape.api.Connector:
            connector = ConnectorFormatValue.from_shape_api(shape.api)
        else:
            connector = None

        return cls(
            box=shape.box,
            id=shape.id,
            line=NaiveLineFormatStyle.from_entity(shape.line),
            connector=connector,
            zorder=shape.api.ZOrderPosition,
        )

    def _common_apply(self, shape: Shape):
        shape.box = self.box
        self.line.apply(shape.line)

    def create_entity(self, context: Context) -> Shape:
        shapes_api = context.shapes.api
        box = self.box
        shape_api = shapes_api.AddLine(
            box.left,
            box.top,
            box.left + box.width,
            box.top + box.height,
        )
        shape = Shape(shape_api)
        self._common_apply(shape)
        if self.connector:
            self.connector.apply(shape, context.shape_id_mapping, context.shapes)
        return shape

    def apply(self, entity: Shape) -> Shape:
        shape = entity
        self._common_apply(shape)
        # NOTE: In case of `apply`, the mapping of id is identical.
        id_mapping = {shape.id: shape.id for shape in shape.slide.shapes}
        if self.connector:
            self.connector.apply(shape, id_mapping, shape.slide.shapes)
        return shape


class FallbackShapeStateModel(FrozenBaseShapeStateModel):
    type: int

    @classmethod
    def from_entity(cls, entity: Shape) -> Self:
        shape = entity
        return cls(box=shape.box,
                   id=shape.id,
                   type=shape.api.Type,
                   zorder=shape.api.ZOrderPosition,
                   )

    def create_entity(self, context: Context) -> Shape:
        shapes = context.shapes
        shape = shapes.add(1)
        shape.text = f"Created, but `{self.type}` cannnot be handled."
        self.apply(shape)
        return shape

    def apply(self, entity: Shape) -> Shape:
        shape = entity
        shape.box = self.box
        return shape

ShapeStateModelValidElements = AutoShapeStateModel | TableShapeStateModel | PictureShapeStateModel | TextBoxShapeStateModel | LineShapeStateModel
ShapeStateModelElements = ShapeStateModelValidElements | FallbackShapeStateModel
