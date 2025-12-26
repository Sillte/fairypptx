
from fairypptx.core.utils import get_discriminator_mapping
from fairypptx.states.context import Context
from fairypptx.shape import Shape
from fairypptx.box import Box 
from pydantic import Field
from typing import Annotated, Self
from fairypptx.states.models import FrozenBaseStateModel
from fairypptx.states.shape.composites import ShapeStateModelComposites
from fairypptx.states.shape.elements import FallbackShapeStateModel
from fairypptx.states.shape.elements import ShapeStateModelValidElements, FallbackShapeStateModel


ShapeStateModelImpl = Annotated[ShapeStateModelValidElements | ShapeStateModelComposites, Field(discriminator="type")]


class ShapeStateModel(FrozenBaseStateModel):
    impl: Annotated[ShapeStateModelImpl | FallbackShapeStateModel, Field(description="Based on `Type`, appropriate class is selected.")]

    @property
    def zorder(self) -> int:
        return self.impl.zorder

    def create_entity(self, context: Context) -> Shape:
        return self.impl.create_entity(context)

    @classmethod
    def from_entity(cls, entity: Shape) -> Self:
        cls_mapping = get_discriminator_mapping(ShapeStateModelImpl, "type")
        klass = cls_mapping.get(entity.api.Type)
        if klass:
            impl = klass.from_entity(entity)
        else:
            impl = FallbackShapeStateModel.from_entity(entity)
        return cls(impl=impl, id=impl.id)

    def apply(self, entity: Shape) -> Shape:
        return self.impl.apply(entity)

    @property
    def box(self) -> Box:
        return self.impl.box


if __name__ == "__main__":
    from fairypptx import Shape
    shape = Shape()
    model = ShapeStateModel.from_entity(shape)
    model.create_entity(Context())
    print(model)
