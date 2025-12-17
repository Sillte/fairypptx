import io
import base64
from PIL import Image

from fairypptx import registry_utils
from fairypptx import constants
from fairypptx.core.utils import get_discriminator_mapping
from fairypptx.states.models import BaseStateModel, FrozenBaseStateModel
from fairypptx.states.context import Context
from fairypptx.constants import msoTrue, msoFalse
from fairypptx.shape import Shape, TableShape, GroupShape
from fairypptx.shape_range import ShapeRange
from fairypptx.box import Box 
from pydantic import Field, Base64Bytes
from typing import Annotated, Self, Literal, cast, Sequence
from fairypptx.styles.line_format import NaiveLineFormatStyle
from fairypptx.styles.fill_format import NaiveFillFormatStyle
from fairypptx.states.table import TableValueModel
from fairypptx.states.text_frame import TextFrameValueModel
from fairypptx.table import Table
from fairypptx.enums import MsoShapeType

class AutoShapeStateModel(FrozenBaseStateModel):
    type: Annotated[Literal[MsoShapeType.AutoShape], Field(description="Type of Shape")] = MsoShapeType.AutoShape
    box: Annotated[Box, Field(description="Represents the position of the shape")]  # (Note that this is jsonable).
    auto_shape_type: Annotated[int, Field(description="Represents MSOAutoShapeType.")]
    line: Annotated[NaiveLineFormatStyle, Field(description="Represents the format of `Line` around the Shape.")]
    fill: Annotated[NaiveFillFormatStyle, Field(description="Represents the format of `Fill` of the Shape.")]
    text_frame: Annotated[TextFrameValueModel, Field(description="Represents the texts of the Shape.")]
    zorder: Annotated[int, Field(description="The value of Zorder")]

    @classmethod
    def from_entity(cls, entity: Shape) -> Self:
        shape = entity
        return cls(box=shape.box,
                   id=shape.id, 
                   auto_shape_type=shape.api.AutoShapeType, 
                   line=NaiveLineFormatStyle.from_entity(shape.line), 
                   fill=NaiveFillFormatStyle.from_entity(shape.fill),
                   text_frame=TextFrameValueModel.from_object(shape.text_frame),
                   zorder=shape.api.ZOrderPosition, 
                   )

    def create_entity(self, context: Context) -> Shape: 
        shapes = context.shapes
        shape = shapes.add(shape_type=self.auto_shape_type)
        self.apply(shape)
        return shape

    def apply(self, entity: Shape) -> Shape:
        shape = entity
        shape.box = self.box
        shape.api.AutoShapeType = self.auto_shape_type
        self.line.apply(shape.line)
        self.fill.apply(shape.fill)
        self.text_frame.apply(shape.text_frame)
        return shape

class TableShapeStateModel(FrozenBaseStateModel):
    type: Annotated[Literal[MsoShapeType.Table], Field(description="Type of Shape")] = MsoShapeType.Table
    box: Annotated[Box, Field(description="Represents the position of the shape")]  # (Note that this is jsonable).
    table: Annotated[TableValueModel, Field(description="Table of the Shape")]
    zorder: Annotated[int, Field(description="The value of Zorder")]
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



class FallbackShapeStateModel(FrozenBaseStateModel):
    box: Annotated[Box, Field(description="Represents the position of the shape")] 
    zorder: Annotated[int, Field(description="The value of Zorder")]
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

class PictureShapeStateModel(BaseStateModel):
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


ShapeStateModelImpl = AutoShapeStateModel | TableShapeStateModel | PictureShapeStateModel


class GroupShapeStateModel(BaseStateModel):
    """
    Note: Only this class, `apply` is not used at `create_entity`,
    since the mapping of `children` is necessary.
    """
    type: Annotated[Literal[MsoShapeType.Group], Field(description="Type of Shape")] = MsoShapeType.Group
    box: Annotated[Box, Field(description="Represents the position of the shape")]  # (Note that this is jsonable).
    children: Annotated[Sequence[ShapeStateModelImpl | FallbackShapeStateModel], Field(description="Shape Children")]
    zorder: Annotated[int, Field(description="The value of Zorder")]

    @classmethod
    def from_entity(cls, entity: Shape) -> Self:
        shape = cast(GroupShape, entity)
        cls_mapping = get_discriminator_mapping(ShapeStateModelImpl, "type")
        impl_children: list[ShapeStateModelImpl | FallbackShapeStateModel] = []
        for child in shape.children:
            klass = cls_mapping.get(child.api.Type)
            if klass:
                impl = klass.from_entity(child)
            else:
                impl = FallbackShapeStateModel.from_entity(child)
            impl_children.append(impl)
        return cls(box=shape.box,
                   id=shape.id, 
                   children=impl_children,
                   zorder=shape.api.ZOrderPosition,
                   )

    def create_entity(self, context: Context) -> Shape:
        children_shapes = []
        for child in self.children:
            children_shapes.append(child.create_entity(context))
        shape = ShapeRange(children_shapes).group()
        # Note: in this function `apply` is unavailable, since `id` is different from the created one. 
        shape.box = self.box
        return shape

 
    def apply(self, entity: Shape) -> Shape:
        shape = cast(GroupShape, entity)
        shape.box = self.box
        id_to_child_model = {child.id: child for child in self.children}
        id_to_child_entity = {child.id: child for child in shape.children}
        keys = id_to_child_model.keys() & id_to_child_entity.keys()
        if len(keys) != len(id_to_child_entity): 
            print("Inconsistency of `id` occurs in `GroupShapeStateModel`")

        if len(keys) != len(id_to_child_model): 
            print("Inconsistency of `id` occurs in `GroupShapeStateModel`")

        keys = sorted(keys, key=lambda s: id_to_child_model[s].zorder)


        for id_key in keys:
            child_model = id_to_child_model[id_key]
            child_entity = id_to_child_entity[id_key]
            child_model.apply(child_entity)

        # 2. ZOrder を最後に調整
        ordered_ids = sorted(
            id_to_child_model.keys(),
            key=lambda id_: id_to_child_model[id_].zorder,
            reverse=True
        )



        for id_ in ordered_ids:
            shape = id_to_child_entity[id_]
            shape.api.ZOrder(constants.msoBringToFront)
        return cast(Shape, shape)



class ShapeStateModel(BaseStateModel):
    impl: Annotated[ShapeStateModelImpl | GroupShapeStateModel | FallbackShapeStateModel, Field(description="Based on `Type`, appropriate class is selected.")]

    def create_entity(self, context: Context) -> Shape:
        return self.impl.create_entity(context)


    @classmethod
    def from_entity(cls, entity: Shape) -> Self:
        cls_mapping = get_discriminator_mapping(ShapeStateModelImpl, "type")
        klass = cls_mapping.get(entity.api.Type)
        if klass:
            impl = klass.from_entity(entity)
        elif entity.api.Type == MsoShapeType.Group:
            impl = GroupShapeStateModel.from_entity(entity)
        else:
            impl = FallbackShapeStateModel.from_entity(entity)
        return cls(impl=impl, id=impl.id)

    def apply(self, entity: Shape) -> Shape:
        return self.impl.apply(entity)


if __name__ == "__main__":
    from fairypptx import Shape
    shape_state = ShapeStateModel.from_entity(Shape())
    print(shape_state)

