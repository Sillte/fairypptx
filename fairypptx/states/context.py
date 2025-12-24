from fairypptx.presentation import Presentation
from fairypptx.slide import Slide
from fairypptx.shapes import Shapes
from typing import Mapping


class Context:
    def __init__(
        self,
        slide: Slide | None = None,
        presentation: Presentation | None = None,
    ):
        if slide is None:
            slide = Slide()
        if presentation is None:
            presentation = slide.presentation
        self.slide = slide
        self.presentation = presentation

        self._shape_id_mapping: dict[int, int] = dict()

    @property
    def shapes(self) -> Shapes:
        return self.slide.shapes

    @property
    def shape_id_mapping(self) -> Mapping[int, int]:
        """Mapping of the `Id` of `ShapeStateModel` and `Shape`.
        Basically, this is intended to use in `create_entity`. 
        In most cases, `Id` of `Shape` on the slides are different from those of `ShapeStateModel`. 
        """
        return self._shape_id_mapping

    def update_id_mapping(self, model_shape_id: int, entity_shape_id: int): 
        self._shape_id_mapping[model_shape_id] = entity_shape_id

