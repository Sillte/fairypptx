from typing import Self, Sequence, Annotated
from pydantic import Field
from fairypptx.states.models import FrozenBaseStateModel
from fairypptx.states.context import Context
from fairypptx.states.shape import ShapeStateModel
from fairypptx.states.text_frame import TextFrameValueModel
from fairypptx import constants
from fairypptx.slide import Slide

from fairypptx import Shape
from fairypptx.states.shape import ShapeStateModel
from fairypptx.states.context import Context
from pydantic import BaseModel
from fairypptx.enums import MsoShapeType



class SlideLayoutShapesValue(BaseModel):
    """It contains `Shapes`,
    whose type is not PlaceHolder(14).
    """

    layout_index: Annotated[int, Field(description="Index of Layout, 32 is not useful number.")]
    layout_name: Annotated[str, Field(description="The name of `CustomLayout`")]
    design_name: Annotated[str, Field(description="The name of `Design`")]

    # Below is for the reconstruction when the model is applied or created to the diffrent presentation.
    layout_shapes: Annotated[
        Sequence[ShapeStateModel],
        Field(
            description="Shapes of Layout, if the exact recover is necessary, it is used."
        ),
    ]
    master_shapes: Annotated[
        Sequence[ShapeStateModel],
        Field(
            description="Shapes of SlideMaster, if the exact recover is necessary, it is used."
        ),
    ]

    @classmethod
    def from_slide(cls, slide: Slide) -> Self:
        layout_shapes = cls._to_layout_states(slide)
        master_shapes = cls._to_master_states(slide)
        layout_name = slide.api.CustomLayout.Name
        design_name = slide.api.Design.Name
        layout_index = slide.api.Layout
        return cls(
            layout_index=layout_index,
            layout_name=layout_name,
            design_name=design_name,
            layout_shapes=layout_shapes,
            master_shapes=master_shapes,
        )

    def create_slide(self, context: Context) -> Slide:
        slide = context.presentation.slides.add(index=-1, layout=12)
        key_to_api = {
            (design_api.Name, layout_api.Name): layout_api
            for design_api in context.presentation.api.Designs
            for layout_api in design_api.SlideMaster.CustomLayouts
        }
        api = key_to_api.get((self.design_name, self.layout_name))
        if api:
            slide.api.CustomLayout = api
        else:
            print("CustomLayout is not found, so the workaround of direct creation is performed.")
            for model in self.layout_shapes:
                model.create_entity(Context(slide=slide))
            for model in self.master_shapes:
                model.create_entity(Context(slide=slide))
        return slide

    @classmethod
    def _to_layout_states(cls, slide: Slide) -> Sequence[ShapeStateModel]:
        layout = slide.api.CustomLayout
        shapes = [
            Shape(shp_api)
            for shp_api in layout.Shapes
            if shp_api.Type != MsoShapeType.PlaceHolder
        ]
        return [ShapeStateModel.from_entity(shape) for shape in shapes]

    @classmethod
    def _to_master_states(cls, slide: Slide) -> Sequence[ShapeStateModel]:
        master = slide.api.Master
        shapes = [
            Shape(shp_api)
            for shp_api in master.Shapes
            if shp_api.Type != MsoShapeType.PlaceHolder
        ]
        return [ShapeStateModel.from_entity(shape) for shape in shapes]


class SlideStateModel(FrozenBaseStateModel):
    shapes: Annotated[
        Sequence[ShapeStateModel], Field(description="StateModels regarding Shape")
    ]
    layout: Annotated[
        SlideLayoutShapesValue,
        Field(description="Non-PlaceHolder Shapes of Slide Layout or specific number"),
    ]

    note_text_frame: Annotated[
        TextFrameValueModel | None, Field(description="TextFrame for Notes of Slide.")
    ] = None

    # The below is Read only attributes...
    # For units of `Size` refer to `Slide.size` (Pt)
    slide_size: Annotated[
        tuple[float, float], Field(description="Any value of PpSlideLayout.")
    ]

    def create_entity(self, context: Context) -> Slide:
        slide = self.layout.create_slide(context) 
        place_holder_shapes = [
            shape
            for shape in slide.shapes
            if shape.api.Type == MsoShapeType.PlaceHolder
        ]
        for shape in place_holder_shapes:
            shape.api.Delete()

        for shape_state in sorted(self.shapes, key=lambda shape: shape.zorder):
            shape_state.create_entity(Context(slide=slide))
        self._common_setting(slide)
        return slide

    @classmethod
    def from_entity(cls, entity: Slide) -> Self:
        slide = entity
        shapes = [ShapeStateModel.from_entity(shape) for shape in slide.shapes]
        note_text_frame = slide.note_text_frame
        note_text_frame_model = (
            TextFrameValueModel.from_object(note_text_frame)
            if note_text_frame.text_range.text
            else None
        )
        return cls(
            id=slide.id,
            note_text_frame=note_text_frame_model,
            shapes=shapes,
            slide_size=slide.size,
            layout=SlideLayoutShapesValue.from_slide(slide),
        )

    def _common_setting(self, slide: Slide) -> None:
        if self.note_text_frame:
            self.note_text_frame.apply(slide.note_text_frame)

        if self.slide_size != slide.size:
            print(
                f"There is difference of size: `{self.slide_size=}` vs. `{slide.size=}`"
            )

    def apply(self, entity: Slide) -> Slide:
        """
        Apply only to shapes with the same id.
        Creation / deletion of shapes is not handled here.
        """
        slide = entity
        self._common_setting(slide)

        # NOTE: Here, `self.layout` is NOT assigned to the slide.
        # Reason:
        # 1. Via this operation, `PlaceHolder` appears.
        # 2. This is unncessary, since when only `apply` is used
        # the equivalent of `slide` is assured.

        # Shapes related.
        id_to_model = {s.id: s for s in self.shapes}
        id_to_entity = {s.id: s for s in slide.shapes}

        common = id_to_model.keys() & id_to_entity.keys()
        for k in common:
            id_to_model[k].apply(id_to_entity[k])
        if id_to_model.keys() != id_to_entity.keys():
            print(
                "Inconsitency happens in `apply` in `SlideStateModel`.",
                "(BaseModelIds)",
                id_to_model.keys(),
                "(EntityIds)",
                id_to_entity.keys(),
            )

        # The large value of Zorder is the front.
        ordered_ids = sorted(
            common, key=lambda id_: id_to_model[id_].zorder, reverse=False
        )

        for id_ in ordered_ids:
            shape = id_to_entity[id_]
            shape.api.ZOrder(constants.msoBringToFront)
        return slide


if __name__ == "__main__":
    slide = Slide()
    model = SlideStateModel.from_entity(slide)
    model.create_entity(Context())
