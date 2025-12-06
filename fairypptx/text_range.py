import itertools
from typing import Sequence, TYPE_CHECKING, Self, Literal


from collections.abc import Sequence

from fairypptx import Application
from fairypptx import object_utils
from fairypptx.object_utils import is_object, upstream
from fairypptx import registry_utils
from fairypptx.core.application import Application
from fairypptx import constants
from fairypptx._text import Text, Font, ParagraphFormat
from fairypptx._text.textrange_stylist import ParagraphTextRangeStylist

from fairypptx.text_frame import TextFrame  # NOQA
from fairypptx.core.resolvers import resolve_text_range
from fairypptx.core.types import COMObject

if TYPE_CHECKING:
    from fairypptx.shape import Shape

class TextRange:
    def __init__(self, arg=None) -> None:
        self.app = Application()
        self._api = resolve_text_range(arg)

    @property
    def api(self) -> COMObject:
        return self._api

    @property
    def api2(self) -> COMObject:
        start = self._api.Start
        length = self._api.Length
        shape_api = upstream(self._api, "Shape")
        return shape_api.TextFrame2.TextRange.GetCharacters(start, length)

    @property
    def shape(self) -> "Shape":
        from fairypptx.shape import Shape
        return Shape(upstream(self.api, "Shape"))

    @property
    def characters(self) -> Sequence["TextRange"]:
        return [TextRange(elem) for elem in self.api.Characters()]

    @property
    def words(self) -> Sequence["TextRange"]:
        return [TextRange(elem) for elem in self.api.Words()]

    @property
    def lines(self) -> Sequence["TextRange"]:
        return [TextRange(elem) for elem in self.api.Lines()]

    @property
    def sentences(self) -> Sequence["TextRange"]:
        return [TextRange(elem) for elem in self.api.Sentences()]

    @property
    def paragraphs(self) -> Sequence["TextRange"]:
        return [TextRange(elem) for elem in self.api.Paragraphs()]

    @property
    def runs(self) -> Sequence["TextRange"]:
        # (2022/02/08): Experimentally, I feel it is better that `runs` are separated at `paragraphs` 
        # Since the modification of `run` affects unintuitive. 
        # This phenomena was seen revising `FontResizer`.
        return [TextRange(elem) for para in self.paragraphs for elem in para.api.Runs()]
        #return [TextRange(elem) for elem in self.api.Runs()]

    @property
    def root(self) -> "TextRange":
        """Return the entire `TextRange`.
        """
        try:
            textframe_api = upstream(self.api, "TextFrame")
        except Exception as e:
            print(e)
            raise e
        return TextRange(textframe_api)

    @property
    def paragraph_index(self) -> int:
        """Return the index of `Paragraph`.
        where  the `Start` of `self` is included.

        [IMPORTANT CAUTION]
        Notice that if the `Delete` or `Insert`
        is performed, then this index may chagnes.  
        You should be careful.
        """
        pivot = self.api.Start
        root = self.root
        for index, para in enumerate(root.paragraphs):
            start, length = para.api.Start, para.api.Length
            if start <= pivot < start + length:
                return index
        raise RuntimeError("Implementation Bug.")

    def insert(self, text: str, mode: Literal["after", "before"]="after") -> "TextRange":
        """Insert the text.
        [TODO] Survey the specification.
        """
        assert mode in {"after", "before"}
        insert_funcs = dict()
        insert_funcs["after"] = self.api.InsertAfter
        insert_funcs["before"] = self.api.InsertBefore
        insert_func = insert_funcs[mode]

        api_object = insert_func(str(text))
        tr = TextRange(api_object)
        tr.text  = text
        return tr

    @property
    def n_tail_newlines(self) -> int:
        """Return the number of consecutive newlines 
        at the tail of `paragraph`, including itself.
        """
        CR_CHARS = {"\r", "\013"}
        text = self.text
        root = self.root
        start, length = self.api.Start, self.api.Length
        n_inner = len(list(itertools.takewhile(lambda t: t in CR_CHARS, reversed(text))))
        next_start = start + length 
        n_outer = 0 
        while next_start + n_outer <= root.api.Length:
            if root.api.Characters(next_start + n_outer, 1).Text not in CR_CHARS:
                break
            n_outer += 1
        return n_inner + n_outer


    @property
    def n_head_newlines(self) -> int:
        """Return the number of consecutive newlines 
        at the head of `paragraph`, including itself.
        """
        CR_CHARS = {"\r", "\013"}
        text = self.text
        root = self.root
        start, length = self.api.Start, self.api.Length
        n_inner = len(list(itertools.takewhile(lambda t: t in CR_CHARS, text)))
        next_start = start - 1
        n_outer = 0 
        while 1 <= next_start - n_outer:
            if root.api.Characters(next_start - n_outer, 1).Text not in CR_CHARS:
                break
            n_outer += 1
        return n_inner + n_outer


    def set_tail_newlines(self, n_newlines: int =1) -> None:
        """Set the `tail` of `newlines`. 
        [IMPORTANT] If you use this func, 
        `paragraphs` may break.
        """
        # [TODO] For this restriction, We have to consider carefully..  
        if not self.text.strip("\r\013"):
            raise NotImplementedError("Currently, this is not expected for empty `TextRange`.")
        n_current = self.n_tail_newlines
        if n_current == n_newlines:
            return 
        elif n_current < n_newlines:
            diff = n_newlines - n_current
            self.api.InsertAfter("\r" * diff)
        else:
            diff = n_current - n_newlines
            start, length = self.api.Start, self.api.Length
            pivot = start + length - 1
            while 0 <= pivot and self.root.api.Characters(pivot, 1).Text in ["\r", "\013"]:
                pivot -= 1
            text = self.root.api.Characters(pivot + 1, diff).Text
            assert all(c in {"\r", "\013"} for c in text), "set_tail_newlines"
            self.root.api.Characters(pivot + 1, diff).Delete()

    def set_head_newlines(self, n_newlines: int =1) -> None:
        """Set the `head` of `newlines`. 
        [IMPORTANT] If you use this func, 
        `paragraphs` may break.
        """
        # [TODO] For this restriction, We have to consider carefully..  
        if not self.text.strip("\r\013"):
            raise NotImplementedError("Currently, this is not expected for empty `TextRange`.")
        n_current = self.n_head_newlines
        if n_current == n_newlines:
            return 
        elif n_current < n_newlines:
            diff = n_newlines - n_current
            self.api.InsertBefore("\r" * diff)
        else:
            # Is it truly all right? 
            # See `set_tail_newlines`.
            diff = n_current - n_newlines
            start, length = self.api.Start, self.api.Length
            self.root.api.Characters(start - diff, diff).Delete()


    @property
    def text(self):
        return Text(self)

    @text.setter
    def text(self, arg):
        text = Text(arg)
        self.api.Text = str(text)
        self.font = text.font
        self.paragraphformat = text.paragraphformat

    @property
    def font(self):
        return Font(self.api.Font)

    @font.setter
    def font(self, param):
        for key, value in param.items():
            object_utils.setattr(self.api.Font, key, value)

    @property
    def paragraphformat(self):
        result =  ParagraphFormat(self.api.ParagraphFormat)
        return result

    @paragraphformat.setter
    def paragraphformat(self, param):
        paragraph_format = ParagraphFormat(param)
        paragraph_format.apply(self.api.ParagraphFormat)

    
    def itemize(self):
        for elem in self.paragraphs:
            elem.api.ParagraphFormat.Bullet.Visible = constants.msoTrue
            elem.api.ParagraphFormat.Bullet.Type = constants.ppBulletUnnumbered

    def find(self, target : str):
        """Return `list` of `TextRange` whose text is `target`. 
        """
        result = []
        start = 0
        while True:
            cand = self.api.Find(target, start, constants.msoTrue)
            if cand is None:
                break
            start = cand.Start + cand.Length - 1 
            result.append(TextRange(cand))
        return result


    def register(self, key, disk=True):
        """ Currently, depending of Paragraphs,
        Style specification rule is ambiguous. 
        Here, (IndentLevel, #paragraphs)'s format is stored.
        Well, then, I wonder whether other mode is introduced or not.

        """
        formatter = ParagraphTextRangeStylist(self)
        registry_utils.register("TextRange", key, formatter, extension=".pkl", disk=disk)

    def like(self, style):
        if isinstance(style, str):
            formatter = registry_utils.fetch("TextRange", style)
            formatter(self)
            return self
        else:
            raise ValueError("Cannot handle, yet.")

    @classmethod
    def make(cls, arg):
        from fairypptx.shape import Shape
        shape = Shape.make(constants.msoShapeRectangle)
        shape.textrange = arg
        return shape.textrange

    @classmethod
    def make_itemization(cls, arg, format=None):
        assert format is None, "Current Implementation"

        """ [TODO]: I'd like a (crude) markdown conversion?
        """
        from fairypptx import Shape
        assert isinstance(arg, Sequence), "Current Implemenation"
        assert all(isinstance(elem, str) for elem in arg), "Current Implementation"
        shape = Shape.make(constants.msoShapeRectangle)
        shape.api.TextFrame.TextRange.Text = "\r".join(arg)
        tr = TextRange(shape)
        tr.api.ParagraphFormat.Bullet.Visible = True
        tr.api.ParagraphFormat.Bullet.Type = constants.ppBulletUnnumbered

        # Itemization's normal display.
        tr.api.ParagraphFormat.Alignment = constants.ppAlignLeft
        return tr

# For structure hierarchy.
from fairypptx._text.editor import DefaultEditor 
from fairypptx._text.editor import FontResizer 
