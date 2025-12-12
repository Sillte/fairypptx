"""There are codes related to modify `TextRange`.

"""
from fairypptx.text_range import TextRange

import itertools
from typing import Literal

from fairypptx.text_range import TextRange



class TextRangeEditor:
    def __init__(self,  text_range: TextRange):
        self.text_range = text_range

    def insert(self, text: str, mode: Literal["after", "before"]="after") -> "TextRange":
        """Insert the text.
        [TODO] Survey the specification.
        """
        assert mode in {"after", "before"}
        insert_funcs = dict()
        insert_funcs["after"] = self.text_range.api.InsertAfter
        insert_funcs["before"] = self.text_range.api.InsertBefore
        insert_func = insert_funcs[mode]

        api_object = insert_func(str(text))
        tr = TextRange(api_object)
        tr.api.Text  = text
        return tr

    @property
    def n_tail_newlines(self) -> int:
        """Return the number of consecutive newlines 
        at the tail of `paragraph`, including itself.
        """
        api = self.text_range.api
        CR_CHARS = {"\r", "\013"}
        text = self.text_range.text
        root = self.text_range.root

        start, length = api.Start, api.Length
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
        api = self.text_range.api
        CR_CHARS = {"\r", "\013"}
        text = self.text_range.text
        root = self.text_range.root
        start, length = api.Start, api.Length
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
        if not self.text_range.text.strip("\r\013"):
            raise NotImplementedError("Currently, this is not expected for empty `TextRange`.")
        n_current = self.n_tail_newlines
        if n_current == n_newlines:
            return 
        elif n_current < n_newlines:
            diff = n_newlines - n_current
            self.text_range.api.InsertAfter("\r" * diff)
        else:
            diff = n_current - n_newlines
            start, length = self.text_range.api.Start, self.text_range.api.Length
            pivot = start + length - 1
            while 0 <= pivot and self.text_range.root.api.Characters(pivot, 1).Text in ["\r", "\013"]:
                pivot -= 1
            text = self.text_range.root.api.Characters(pivot + 1, diff).Text
            assert all(c in {"\r", "\013"} for c in text), "set_tail_newlines"
            self.text_range.root.api.Characters(pivot + 1, diff).Delete()

    def set_head_newlines(self, n_newlines: int =1) -> None:
        """Set the `head` of `newlines`. 
        [IMPORTANT] If you use this func, 
        `paragraphs` may break.
        """
        # [TODO] For this restriction, We have to consider carefully..  
        if not self.text_range.text.strip("\r\013"):
            raise NotImplementedError("Currently, this is not expected for empty `TextRange`.")
        n_current = self.n_head_newlines
        if n_current == n_newlines:
            return 
        elif n_current < n_newlines:
            diff = n_newlines - n_current
            self.text_range.api.InsertBefore("\r" * diff)
        else:
            # Is it truly all right? 
            # See `set_tail_newlines`.
            diff = n_current - n_newlines
            start, length = self.text_range.api.Start, self.text_range.api.Length
            self.text_range.api.Characters(start - diff, diff).Delete()
