from pathlib import Path
from fairypptx.core.resolvers import resolve_presentation
from fairypptx.core.types import COMObject, ObjectLike

class Presentation:
    def __init__(self, arg: None | str | Path | ObjectLike = None):
        self._api = resolve_presentation(arg) 

    @property
    def api(self) -> COMObject:
        return self._api

    @property
    def slides(self):
        from fairypptx.slides import Slides
        return Slides(self.api.Slides)

