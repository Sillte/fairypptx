from fairypptx import TextRange
from fairypptx.parts._markdown.jsonast.models import BaseBlock, BaseInlineModel


from contextlib import contextmanager
from dataclasses import dataclass
from typing import Iterator, Mapping, Type, TYPE_CHECKING

if TYPE_CHECKING:
    from fairypptx.parts._markdown.jsonast.editors import BlockEditProtocol, InlineEditProtocol


@dataclass(frozen=True)
class FontState:
    underline: bool = False
    bold: bool = False

    def apply(self, text_range: TextRange) -> None:
        """Only the first target `TextRange` is given.
        When the change of `FontState` is detected.
        """
        text_range.font.underline = self.underline
        text_range.font.bold = self.bold

@dataclass(frozen=True)
class ParagraphState:
    indent_level: int = 1
    bullet_type: int | None = None

    def apply(self, text_range: TextRange) -> None:
        """Only the first target `TextRange` is given.
        When the change of `FontState` is detected.
        """
        text_range.paragraph_format.indent_level = self.indent_level
        text_range.paragraph_format.bullet_type = self.bullet_type


@dataclass
class Context:
    text_range: TextRange
    block_editors: Mapping[Type[BaseBlock], Type["BlockEditProtocol"]]
    inline_editors: Mapping[Type[BaseInlineModel], Type["InlineEditProtocol"]]
    font_state: FontState = FontState()
    paragraph_state = ParagraphState()

    def __post_init__(self):
        if self.font_state is None:
            self.font_state = FontState(bold=False, underline=False)
        self._need_to_apply_font = True

        if self.paragraph_state is None:
            self.paragraph_state = ParagraphState()
        self._need_to_apply_paragraph = True


    def apply_inline(self, inline: BaseInlineModel):
        from fairypptx.parts._markdown.jsonast.editors import FallbackInlineEditor
        function = self._fetch_regisered_instance(inline, self.inline_editors) or FallbackInlineEditor()
        function(inline, self)

    def apply_block(self, block: BaseBlock):
        from fairypptx.parts._markdown.jsonast.editors import FallbackBlockEditor
        function = self._fetch_regisered_instance(block, self.block_editors) or FallbackBlockEditor()
        function(block, self)

    def _fetch_regisered_instance(self, model: BaseBlock | BaseInlineModel, registers: Mapping) ->  "None | BlockEditProtocol | InlineEditProtocol":
        for cls in type(model).mro():
            if cls in registers:
                return registers[cls]()
        return None


    @contextmanager
    def update_font_state(self, font_state: FontState) -> Iterator[None]:
        prev_font_state = self.font_state
        self.font_state = font_state
        self._need_to_apply_font = True
        yield
        self.font_state = prev_font_state
        self._need_to_apply_font = True

    @contextmanager
    def update_paragraph_state(self, paragraph_state: ParagraphState) -> Iterator[None]:
        prev_state = self.paragraph_state
        self.paragraph_state = paragraph_state
        self._need_to_apply_paragraph = True
        yield
        self.paragraph_state = prev_state
        self._need_to_apply_paragraph = True


    def insert_text(self, text: str):
        """Inser the text, if necessary `format` is modified for the generated TextRange.
        """
        text_range = self.text_range
        if self._need_to_apply_font or self._need_to_apply_paragraph:
            start_idx = text_range.total_count
            text_range.insert(text)
            length = text_range.total_count - start_idx
            if length:
                new_text_range = text_range.get_range_from_root(start_idx + 1, length)
                self.paragraph_state.apply(new_text_range)
                self.font_state.apply(new_text_range)
                self._need_to_apply_font = False
                self._need_to_apply_paragraph = False
        else:
            text_range.insert(text)
