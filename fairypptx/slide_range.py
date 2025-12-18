from typing import Sequence, Self, overload
from collections.abc import Sequence as SeqABC
from fairypptx.core.resolvers import resolve_slide_range
from fairypptx.core.types import COMObject
from fairypptx.slide import Slide
from fairypptx.object_utils import is_object

class SlideRange:

    def __init__(self, arg: COMObject | Sequence[COMObject] |
                       Self | Sequence[Slide] | None = None):
        self._slides: list[Slide] = self._solve_slides(arg)


    def __len__(self):
        return len(self._slides)

    def __iter__(self):
        yield from self._slides

    @overload
    def __getitem__(self, key: int) -> "Slide":
        ...

    @overload
    def __getitem__(self, key: slice) -> "SlideRange":
        ...


    def __getitem__(self, key: int | slice) -> "Slide | SlideRange":
        if isinstance(key, int):
            return self._slides[key]
        elif isinstance(key, slice):
            return SlideRange(self._slides[key])
        else:
            raise TypeError(f"Invalid key: {key!r}")

    @property
    def api(self) -> COMObject:
        """Reconstruct COM SlideRange from stored slides."""
        if not self._slides:
            msg = "It is impossible to get COMObject for the empty range."
            raise ValueError(msg)
        slides_api = self._slides[0].api.Parent  # COM Slides
        if is_object(slides_api, "Presentation"):
            slides_api = slides_api.Slides
        indices = [s.api.SlideIndex for s in self._slides]
        return slides_api.Range(indices)


    def _solve_slides(self, arg) -> list[Slide]:
        """Normalize input → list[Slide]"""

        # 1) COM SlideRange
        if is_object(arg, "SlideRange"):
            return [Slide(arg.Item(i + 1)) for i in range(arg.Count)]

        # 2) Python list of Slide
        if isinstance(arg, SeqABC) and not isinstance(arg, (str, bytes)):
            if all(isinstance(s, Slide) for s in arg):
                return list(arg)

            if all(is_object(s, "Slide") for s in arg):
                return [Slide(s) for s in arg]

        # 3) SlideRange (clone)
        if isinstance(arg, SlideRange):
            return list(arg._slides)

        # 4) Fallback: resolve as COM SlideRange
        api = resolve_slide_range(arg)  # must return SlideRange COM
        
        # To be safe.
        if is_object(api, "Slide"):
            return [Slide(api)]
        return self._solve_slides(api)
