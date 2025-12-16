from fairypptx.presentation import Presentation
from fairypptx.slide import Slide
from fairypptx.shapes import Shapes


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

    @property
    def shapes(self) -> Shapes:
        return self.slide.shapes
