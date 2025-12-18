from fairypptx import constants
from fairypptx.slide_range import SlideRange
from fairypptx.core.resolvers import resolve_slides
from fairypptx.core.types import COMObject
from fairypptx.slide import Slide
from typing import overload, Sequence


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

    @overload
    def __getitem__(self, key: int) -> "Slide":
        ...

    @overload
    def __getitem__(self, key: slice) -> "SlideRange":
        ...
                    
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

    def delete_all(self):
        while self.api.Count >= 1:
            self.api.Item(self.api.Count).Delete()

    def swap(self, slide1: Slide, slide2: Slide) -> None:
        """Swap the orders of 2 slide. 
        """
        s_low, s_high = sorted([slide1, slide2], key=lambda slide: slide.index) # index1 < index2. 
        low_index, high_index = s_low.index, s_high.index
        if low_index == high_index:
            return 
        s_high.api.MoveTo(toPos=low_index)
        s_low.api.MoveTo(toPos=high_index)

    def reorder(self, reorder_indices: Sequence[int]) -> None:
        index_to_slide = {slide.index: slide for slide in self}
        orig_indices = [slide.index for slide in self] 
        if set(orig_indices) != set(reorder_indices):
            raise ValueError("Set of Silde.index must be equivalent.")
        assert set(range(1, len(self) + 1)) == set(index_to_slide.keys())

        for i, index in enumerate(reorder_indices, start=1):
            index_to_slide[index].api.MoveTo(toPos=i)



