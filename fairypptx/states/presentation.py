from typing import Self, Sequence, Annotated
from pydantic import Field
from fairypptx.states.models import  FrozenBaseStateModel
from fairypptx.core.resolvers import Application
from fairypptx.states.context import Context
from fairypptx.states.slide import SlideStateModel
from fairypptx.presentation import Presentation

class PresentationStateModel(FrozenBaseStateModel):
    slides: Annotated[Sequence[SlideStateModel], Field(description="Slides of the presentation")]

    @classmethod
    def from_entity(cls, entity: Presentation) -> Self:
        pres = entity
        slides = [SlideStateModel.from_entity(slide) for slide in pres.slides]
        return cls(slides=slides, id=pres.api.Name)

    def create_entity(self, context: Context) -> Presentation:
        pres = Presentation(Application().api.Presentations.Add())
        context = Context(presentation=pres)
        for slide in self.slides:
            slide.create_entity(context)
        return pres

    def apply(self, entity: Presentation) -> Presentation:
        """
        * Apply the format of the slide to the one with the same id. 
        """
        pres = entity
        # Shapes related.
        id_to_model = {s.id: s for s in self.slides}
        id_to_entity = {s.id: s for s in pres.slides}

        common = id_to_model.keys() & id_to_entity.keys()
        for k in common:
            id_to_model[k].apply(id_to_entity[k])
        is_all_slide_match = (id_to_model.keys() == id_to_entity.keys())
        if  is_all_slide_match:
            # If all_slide_match, then reverse is performed.    
            reorder_ids = [s.id for s in self.slides]
            id_to_index = {id: id_to_entity[id].index for id in common}
            reorder_indices = [id_to_index[id] for id in reorder_ids]
            pres.slides.reorder(reorder_indices)
        else:
            print("Inconsitency happens in `apply` in `PresentationStateModel`.")
        return pres
