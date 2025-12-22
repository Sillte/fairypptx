from typing import Self, Sequence, Annotated
from pydantic import Field
from fairypptx.states.models import  FrozenBaseStateModel
from fairypptx.states.context import Context
from fairypptx.states.shape import ShapeStateModel
from fairypptx.states.text_frame import TextFrameValueModel
from fairypptx import constants
from fairypptx.slide import Slide

class SlideStateModel(FrozenBaseStateModel):
    shapes: Annotated[Sequence[ShapeStateModel], Field(description="StateModels regarding Shape")]
    layout: Annotated[int, Field(description="Any value of PpSlideLayout.") ]

    note_text_frame: Annotated[TextFrameValueModel | None, Field(description="TextFrame for Notes of Slide.")] = None

    # The below is Read only attributes...
    # For units of `Size` refer to `Slide.size` (Pt) 
    slide_size: Annotated[tuple[float, float], Field(description="Any value of PpSlideLayout.") ]

    def create_entity(self, context: Context) -> Slide:
        slide = context.presentation.slides.add(layout=self.layout)
        for shape_state in self.shapes:
            shape_state.create_entity(Context(slide=slide))
        self.apply(slide)
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
            slide_size=slide.size
    )
    def apply(self, entity: Slide) -> Slide:
        """
        Apply only to shapes with the same id.
        Creation / deletion of shapes is not handled here.
        """

        slide = entity
        if self.layout != constants.ppLayoutMixed:
            slide.api.Layout = self.layout
        if self.note_text_frame:
            self.note_text_frame.apply(slide.note_text_frame)

        if self.slide_size != slide.size:
            print(f"There is difference of size: `{self.slide_size=}` vs. `{slide.size=}`")

        # Shapes related.
        id_to_model = {s.id: s for s in self.shapes}
        id_to_entity = {s.id: s for s in slide.shapes}

        common = id_to_model.keys() & id_to_entity.keys()
        for k in common:
            id_to_model[k].apply(id_to_entity[k])
        if id_to_model.keys() != id_to_entity.keys():
            print("Inconsitency happens in `apply` in `SlideStateModel`.")

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


