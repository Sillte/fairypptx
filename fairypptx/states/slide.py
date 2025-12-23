from typing import Self, Sequence, Annotated
from pydantic import Field
from fairypptx.states.models import  FrozenBaseStateModel
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
from fairypptx.box import Box

from fairypptx.presentation import Presentation

class SlideLayoutShapesValue(BaseModel):
    """It contains `Shapes`,
    whose type is not PlaceHolder(14).  
    """
    layout_shapes: Sequence[ShapeStateModel]
    master_shapes: Sequence[ShapeStateModel]

    @classmethod
    def from_slide(cls, slide: Slide) -> Self:
        layout_shapes = cls._to_layout_states(slide)
        master_shapes = cls._to_master_states(slide)
        return cls(layout_shapes=layout_shapes, master_shapes=master_shapes, layout=slide.api.Layout)

    def assume_layout(self, pres: Presentation) -> int:
        target = [shape.box for shape in self.layout_shapes]
        master = pres.api.SlideMaster
        for i, layout in enumerate(master.CustomLayouts, 1):
            shapes = [Shape(shp_api) for shp_api in layout.Shapes if shp_api.Type != MsoShapeType.PlaceHolder]
            boxes = [shape.box for shape in shapes]
            if len(target) == len(boxes) and all(Box.intersection_over_union(elem1, elem2) >= 0.9 for elem1, elem2 in zip(target, boxes)
                                                 ):
                print("Found the same layout", i)
                return i
        return 12

    @classmethod
    def _to_layout_states(cls, slide: Slide) -> Sequence[ShapeStateModel]:
        layout = slide.api.CustomLayout
        shapes = [Shape(shp_api) for shp_api in layout.Shapes if shp_api.Type != MsoShapeType.PlaceHolder]
        return [ShapeStateModel.from_entity(shape) for shape in shapes]

    @classmethod
    def _to_master_states(cls, slide: Slide) -> Sequence[ShapeStateModel]:
        master = slide.api.Master
        shapes = [Shape(shp_api) for shp_api in master.Shapes if shp_api.Type != MsoShapeType.PlaceHolder]
        return [ShapeStateModel.from_entity(shape) for shape in shapes]

class SlideStateModel(FrozenBaseStateModel):
    shapes: Annotated[Sequence[ShapeStateModel], Field(description="StateModels regarding Shape")]
    layout: Annotated[int, Field(description="Any value of PpSlideLayout.") ]

    layout_shapes: Annotated[SlideLayoutShapesValue, Field(description="Non-PlaceHolder Shapes of Slide Layout")]

    note_text_frame: Annotated[TextFrameValueModel | None, Field(description="TextFrame for Notes of Slide.")] = None

    # The below is Read only attributes...
    # For units of `Size` refer to `Slide.size` (Pt) 
    slide_size: Annotated[tuple[float, float], Field(description="Any value of PpSlideLayout.") ]

    def create_entity(self, context: Context) -> Slide:
        if self.layout == constants.ppLayoutCustom:
            layout = self.layout_shapes.assume_layout(context.presentation)
        else:
            layout = self.layout
        slide = context.presentation.slides.add(layout=layout)
        place_holder_shapes = [shape for shape in slide.shapes if shape.api.Type == MsoShapeType.PlaceHolder]
        for shape in place_holder_shapes:
            shape.api.Delete()

        for shape_state in sorted(self.shapes, key=lambda shape: shape.zorder):
            shape_state.create_entity(Context(slide=slide))
        self._common_setting(slide)
        return slide

    @classmethod
    def from_entity(cls, entity: Slide) -> Self:
        slide = entity
        shapes = [
            ShapeStateModel.from_entity(shape)
            for shape in slide.shapes
        ]
        note_text_frame = slide.note_text_frame
        note_text_frame_model = TextFrameValueModel.from_object(note_text_frame) if note_text_frame.text_range.text else None
        return cls(
            id=slide.id,
            layout=slide.api.Layout,
            note_text_frame=note_text_frame_model,
            shapes=shapes,
            slide_size=slide.size,
            layout_shapes=SlideLayoutShapesValue.from_slide(slide)
    )


    def _common_setting(self, slide: Slide) -> None:
        if self.note_text_frame:
            self.note_text_frame.apply(slide.note_text_frame)

        if self.slide_size != slide.size:
            print(f"There is difference of size: `{self.slide_size=}` vs. `{slide.size=}`")


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
            print("Inconsitency happens in `apply` in `SlideStateModel`.",
                  "(BaseModelIds)", id_to_model.keys(),
                  "(EntityIds)", id_to_entity.keys())

        # The large value of Zorder is the front.
        ordered_ids = sorted(
            common,
            key=lambda id_: id_to_model[id_].zorder,
            reverse=False
        )

        for id_ in ordered_ids:
            shape = id_to_entity[id_]
            shape.api.ZOrder(constants.msoBringToFront)
        return slide


if __name__ == "__main__":
    slide = Slide()
    model = SlideStateModel.from_entity(slide)
    model.create_entity(Context())


