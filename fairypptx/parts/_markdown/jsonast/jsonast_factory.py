import subprocess
from itertools import groupby
from fairypptx.text_range import TextRange

from fairypptx.parts.markdown import Markdown
import json
from fairypptx.parts._markdown.jsonast.context import ParagraphState
from fairypptx.parts._markdown.jsonast.utils import to_jsonast
from fairypptx.parts._markdown.jsonast.models import (
    PandocJsonAst,
    ParaBlock,
    StrInline,
    BlockElement,
    InlineElement,
    OrderedList,
    BulletList,
    PlainBlock,
)

from fairypptx import constants

from typing import Mapping, Sequence, Self
from dataclasses import dataclass


def run_to_inline(run: TextRange) -> InlineElement:
    # [TODO]: We have to add the `Inline`, based on `Font`.
    return StrInline(c=run.text)


def paragraph_to_inlines(paragraph: TextRange) -> list[InlineElement]:
    """Note that"""
    inlines = [run_to_inline(run) for run in paragraph.runs]
    return inlines


@dataclass(frozen=True)
class ParaBlockItem:
    paragraphs: Sequence[TextRange]

    def to_pandoc_model(self) -> BlockElement:
        inlines = sum([paragraph_to_inlines(paragraph) for paragraph in self.paragraphs], [])
        return ParaBlock(c=inlines)


@dataclass(frozen=True)
class BulletBlockItem:
    paragraphs: Sequence[TextRange]

    def to_pandoc_model(self) -> BlockElement:
        blocks = self._split_into_items()
        elements = []
        for block in blocks:
            sub_element = [PlainBlock(c=paragraph_to_inlines(block.paragraphs[0]))]
            remain_paragraphs = block.paragraphs[1:]
            if remain_paragraphs:
                remain_blocks = make_block_collection(remain_paragraphs)
                sub_element += [factor.to_pandoc_model() for factor in remain_blocks]
            elements.append(sub_element)
        return self._to_block_element(elements)

    def _to_block_element(self, elements: Sequence[Sequence[BlockElement]]) -> BlockElement:
        bullet_type = self.paragraphs[0].paragraph_format.bullet_type
        if bullet_type == constants.ppBulletNumbered:
            return OrderedList(c=((1, {"t": "Decimal"}, {"t": "Period"}), elements))
        elif bullet_type == constants.ppBulletUnnumbered:
            return BulletList(c=elements)
        elif not bullet_type: 
            raise RuntimeError("Implementation error.")
        else:
            # Fallback
            return BulletList(c=elements)

    def _split_by_same_level(self) -> list[list[TextRange]]:
        paragraphs = self.paragraphs
        states = [ParagraphState.from_text_range(p) for p in paragraphs]

        def same_level(i: int, j: int) -> bool:
            return states[i].bullet_type == states[j].bullet_type and states[i].indent_level == states[j].indent_level

        result = []
        start = 0
        for i in range(1, len(paragraphs)):
            if same_level(start, i):
                result.append(paragraphs[start:i])
                start = i

        result.append(paragraphs[start:])
        return result

    def _split_into_items(self) -> list["BlockItem"]:
        chunks = self._split_by_same_level()
        return [self._build_item(chunk) for chunk in chunks]

    def _build_item(self, paragraphs: Sequence[TextRange]) -> "BlockItem":
        state = ParagraphState.from_text_range(paragraphs[0])
        if state.bullet_type is None:
            return ParaBlockItem(paragraphs)
        return BulletBlockItem(paragraphs)


BlockItem = ParaBlockItem | BulletBlockItem


def make_block_collection(paragraphs: Sequence[TextRange]) -> Sequence[BlockItem]:
    """Return the division of `paragraphs`.  
    Each BlockItem is expectd to generate one block in the domain of Pandoc.  
    """
    def _split_into_block_spans(
        groups: list[list[TextRange]],
    ) -> list[list[TextRange]]:
        def _find_block_end_index(
            groups: list[list[TextRange]],
            start: int,
        ) -> int:
            def _is_block_boundary(start: ParagraphState, current: ParagraphState) -> bool:
                if start.bullet_type is None:
                    return current.bullet_type != start.bullet_type

                return current.bullet_type != start.bullet_type and current.indent_level <= start.indent_level

            start_state = ParagraphState.from_text_range(groups[start][0])

            for i in range(start + 1, len(groups)):
                state = ParagraphState.from_text_range(groups[i][0])
                if _is_block_boundary(start_state, state):
                    return i - 1
            return len(groups) - 1

        spans = []
        index = 0

        while index < len(groups):
            end = _find_block_end_index(groups, index)
            span = sum(groups[index : end + 1], [])
            spans.append(span)
            index = end + 1
        return spans

    def _build_block_item(paragraphs: Sequence[TextRange]) -> BlockItem:
        state = ParagraphState.from_text_range(paragraphs[0])
        if state.bullet_type is None:
            return ParaBlockItem(paragraphs)
        return BulletBlockItem(paragraphs)

    if not paragraphs:
        return []

    groups = [list(group) for _, group in groupby(paragraphs, key=ParagraphState.from_text_range)]
    spans = _split_into_block_spans(groups)
    return [_build_block_item(span) for span in spans]


def to_json_ast(text_range: TextRange) -> PandocJsonAst:
    collections = make_block_collection(text_range.paragraphs)
    blocks = [elem.to_pandoc_model() for elem in collections]
    return PandocJsonAst(blocks=blocks)


class PandocRenderer:
    def __init__(self, json_ast: PandocJsonAst):
        self.json_ast = json_ast

    def to_markdown(
        self,
    ) -> str:
        ast_dict = self.json_ast.model_dump(by_alias=True)

        proc = subprocess.run(
            ["pandoc", "-f", "json", "-t", "markdown"],
            input=json.dumps(ast_dict),
            text=True,
            capture_output=True,
            check=True,
        )
        return proc.stdout


if __name__ == "__main__":
    SCRIPT = """ 
Hello, world

1. BBBB
    * S1
        * S2
    * S3
2. CCC
""".strip()
    json_ast_data = to_jsonast(SCRIPT)
    print(json_ast_data)
    shape = Markdown.make(SCRIPT).shape

    pandoc_json_ast = to_json_ast(shape.text_range)

    # json_ast = to_json_ast(shape.text_range)
    # pandoc_json_ast = PandocJsonAst.model_validate(json_ast_data)
    # from pprint import pprint
    # pprint(json_ast_data)
    print(PandocRenderer(pandoc_json_ast).to_markdown())
