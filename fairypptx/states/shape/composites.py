from fairypptx import GroupShape, Shape, constants
from fairypptx.box import Box
from fairypptx.core.utils import get_discriminator_mapping
from fairypptx.enums import MsoShapeType
from fairypptx.shape import Shape
from fairypptx.shape_range import ShapeRange
from fairypptx.states.context import Context
from fairypptx.states.models import BaseStateModel
from fairypptx.states.shape.elements import FallbackShapeStateModel
from fairypptx.states.shape.elements import ShapeStateModelElements

from pydantic import Field


from typing import Annotated, Literal, Self, Sequence, cast


class PlaceHolderShapeStateModel(BaseStateModel):
    type: Annotated[Literal[MsoShapeType.PlaceHolder], Field(description="Type of Shape")] = MsoShapeType.PlaceHolder
    impl: Annotated[ShapeStateModelElements, Field(description="Internal ShapeStateModel")]

    @property
    def zorder(self):
        return self.impl.zorder

    @classmethod
    def from_entity(cls, entity: Shape) -> Self:
        shape = entity
        contained_type = shape.api.PlaceholderFormat.ContainedType
        cls_mapping = get_discriminator_mapping(ShapeStateModelElements, "type")
        klass = cls_mapping.get(contained_type, FallbackShapeStateModel)
        return cls(id=shape.id, impl=klass.from_entity(entity))


    def create_entity(self, context: Context) -> Shape:
        print("PlaceHolderShape is unable to be created, so the alternative autoshape is generated")
        return self.impl.create_entity(context)


    def apply(self, entity: Shape) -> Shape:
        shape = entity
        self.impl.apply(shape)
        return shape

    @property
    def box(self) -> Box:
        return self.impl.box


class GroupShapeStateModel(BaseStateModel):
    """
    Note: Only this class, `apply` is not used at `create_entity`,
    since the mapping of `children` is necessary.
    """
    type: Annotated[Literal[MsoShapeType.Group], Field(description="Type of Shape")] = MsoShapeType.Group
    box: Annotated[Box, Field(description="Represents the position of the shape")]  # (Note that this is jsonable).
    children: Annotated[Sequence[ShapeStateModelElements], Field(description="Shape Children")]
    zorder: Annotated[int, Field(description="The value of Zorder")]

    @classmethod
    def from_entity(cls, entity: Shape) -> Self:
        shape = cast(GroupShape, entity)
        cls_mapping = get_discriminator_mapping(ShapeStateModelElements, "type")
        impl_children: list[ShapeStateModelElements] = []
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
        # The large value of Zorder is the front. 
        ordered_ids = sorted(
            id_to_child_model.keys(),
            key=lambda id_: id_to_child_model[id_].zorder,
            reverse=False
        )

        for id_ in ordered_ids:
            shape = id_to_child_entity[id_]
            shape.api.ZOrder(constants.msoBringToFront)
        return cast(Shape, shape)


ShapeStateModelComposites = PlaceHolderShapeStateModel | GroupShapeStateModel
