from fairypptx import constants
from fairypptx.slide_range import SlideRange
from fairypptx.core.resolvers import resolve_slides
from fairypptx.core.types import COMObject
from fairypptx.slide import Slide


class Slides:
    """Slides.
    However, this may not behave like all of the slides 
    is handled this class.   
    It accepts a subset of Slides Object. 

    Note
    ---------------------
    * `Add` / `Delete` operations may break this class.

    """

    def __init__(self, arg=None):
        self._api = resolve_slides(arg)

    @property
    def api(self) -> COMObject:
        return self._api

    def add(self, index: int | None = None, layout: int | None = None) -> "Slide":
        # You may consider the branch processing.
        if index is None:
            index = len(self)
        assert 0 <= index <= len(self)
        if layout is None:
            layout = constants.ppLayoutBlank
        return Slide(self.api.Add(index + 1, layout))

    def __len__(self) -> int:
        return self.api.Count

    def __getitem__(self, key: int | slice) -> "Slide | SlideRange":
        if isinstance(key, int):
            return Slide(self.api.Item(key + 1))
        elif isinstance(key, slice):
            indices = range(*key.indices(self.api.Count))
            dispatch_list = [self.api.Item(i+1) for i in indices]
            return SlideRange(dispatch_list)
        msg = f"`{key}` is unacceptable."
        raise ValueError(msg)

    def __iter__(self):
        for i in range(self.api.Count):
            yield self[i]
