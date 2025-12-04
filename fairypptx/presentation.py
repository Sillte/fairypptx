from fairypptx.core.application import Application
from pathlib import Path
from pywintypes import com_error
from fairypptx.object_utils import is_object
from fairypptx.core.resolvers import resolve_presentation
from fairypptx.core.types import COMObject, ObjectLike

class Presentation:
    def __init__(self, arg: None | str | Path | ObjectLike = None, app=None):
        self._api = resolve_presentation(arg) 

    @property
    def api(self) -> COMObject:
        return self._api

    @property
    def slides(self):
        from fairypptx.slide import Slides
        return Slides(self.api.Slides)

