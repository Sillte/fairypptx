from dataclasses import replace
from typing import Mapping, Protocol, TYPE_CHECKING
from fairypptx.parts._markdown.jsonast.models import BaseBlock, BaseInlineModel, CodeBlock, CodeInline, Header, LineBreakInline, ParaBlock, PlainBlock, SoftBreakInline, SpaceInline, StrInline, StrongInline, LinkInline, OrderedList, BulletList, QuotedInline
from fairypptx.parts._markdown.jsonast.context import Context
from fairypptx import constants


class BlockEditProtocol[T: BaseBlock](Protocol):
    def __call__(self, model: T, context: "Context"):
        ...


BLOCK_EDITORS: dict[type[BaseBlock], type[BlockEditProtocol]] = {}

def get_default_block_editors() -> Mapping[type[BaseBlock], type[BlockEditProtocol]]:
    return dict(BLOCK_EDITORS)


def register_block(block_type: type[BaseBlock]):
    def decorator(editor_cls: type[BlockEditProtocol]):
        BLOCK_EDITORS[block_type] = editor_cls
        return editor_cls
    return decorator


@register_block(ParaBlock)
class ParaEditor(BlockEditProtocol):
    def __call__(self, model: ParaBlock, context: Context):
        for inline in model.inlines:
            context.apply_inline(inline)
        context.insert_text("\r")


@register_block(PlainBlock)
class PlainEditor(BlockEditProtocol):
    def __call__(self, model: PlainBlock, context: Context):
        for inline in model.inlines:
            context.apply_inline(inline)
        context.insert_text("\r")


@register_block(CodeBlock)
class CodeBlockEditor(BlockEditProtocol):
    def __call__(self, model: CodeBlock, context: Context):
        text = model.c[1]
        context.insert_text(text)


@register_block(Header)
class HeaderEditor(BlockEditProtocol):
    def __call__(self, model: Header, context: Context):
        font_state = context.font_state
        assert font_state
        with context.update_font_state(replace(font_state, bold=True, underline=True)):
            for inline in model.inlines:
                context.apply_inline(inline)
            context.insert_text("\r")


@register_block(BulletList)
class BulletListEditor(BlockEditProtocol):
    def __call__(self, model: BulletList, context: Context):
        current = context.paragraph_state
        if current.bullet_type is not None:
            next_indent_level = current.indent_level + 1
        else:
            next_indent_level = current.indent_level
        new_state = replace(current, 
                            indent_level=next_indent_level,
                            bullet_type=constants.ppBulletUnnumbered)
        with context.update_paragraph_state(new_state):
            blocks_list = model.blocks_list
            for blocks in blocks_list:
                for block in blocks:
                    context.apply_block(block)


@register_block(OrderedList)
class OrderedListEditor(BlockEditProtocol):
    def __call__(self, model: OrderedList, context: Context):
        current = context.paragraph_state
        if current.bullet_type is not None:
            next_indent_level = current.indent_level + 1
        else:
            next_indent_level = current.indent_level
        new_state = replace(current, 
                            indent_level=next_indent_level,
                            bullet_type=constants.ppBulletNumbered)
        with context.update_paragraph_state(new_state):
            blocks_list = model.blocks_list
            for blocks in blocks_list:
                for block in blocks:
                    context.apply_block(block)



class FallbackBlockEditor(BlockEditProtocol):
    def __call__(self, model: BaseBlock, context: Context):
        print(f"[WARNING]{model=} cannot be used for translation.\n Confirm the Editor of `{model.__class__}`")


class InlineEditProtocol[T: BaseInlineModel](Protocol):
    def __call__(self, model: T, context: "Context"):
        ...


INLINE_EDITORS: dict[type[BaseInlineModel], type[InlineEditProtocol]] = {}

def get_default_inline_editors() -> Mapping[type[BaseInlineModel], type[InlineEditProtocol]]:
    return dict(INLINE_EDITORS)

def register_inline(block_type: type[BaseInlineModel]):
    def decorator(editor_cls: type[InlineEditProtocol]):
        INLINE_EDITORS[block_type] = editor_cls
        return editor_cls
    return decorator


class FallbackInlineEditor(InlineEditProtocol):
    def __call__(self, model: BaseInlineModel, context: Context):
        print(f"[WARNING]{model=} cannot be used for translation.\n Confirm the Editor of  `{model.__class__}`")


@register_inline(StrInline)
class StrEditor(InlineEditProtocol):
    def __call__(self, model: StrInline, context: Context):
        context.insert_text(model.c)


@register_inline(StrongInline)
class StrongEditor(InlineEditProtocol):
    def __call__(self, model: StrongInline, context: Context):
        font_state = context.font_state
        assert font_state
        with context.update_font_state(replace(font_state, bold=True)):
            for inline in model.inlines:
                context.apply_inline(inline)


@register_inline(SpaceInline)
class SpaceEditor(InlineEditProtocol):
    def __call__(self, _: SpaceInline, context: Context):
        text_range = context.text_range
        text_range.insert(" ")


@register_inline(LineBreakInline)
class LineBreakEditor(InlineEditProtocol):
    def __call__(self, model: LineBreakInline, context: Context):
        context.insert_text("\n")


@register_inline(SoftBreakInline)
class SoftBreakEditor(InlineEditProtocol):
    def __call__(self, model: SoftBreakInline , context: Context):
        context.insert_text("\n")

@register_inline(QuotedInline)
class QuotedInlineEditor(InlineEditProtocol):
    def __call__(self, model: QuotedInline, context: Context):
        quote_type, inlines = model.c
        if quote_type.t == "SingleQuote":
            quote = "'"
        else:
            quote = '"'
        context.insert_text(quote)
        for inline in inlines:
            context.apply_inline(inline)
        context.insert_text(quote)


@register_inline(CodeInline)
class CodeInlineEditor(InlineEditProtocol):
    def __call__(self, model: CodeInline , context: Context):
        text = model.c[1]
        context.insert_text(text)

@register_inline(LinkInline)
class LinkInlineEditor(InlineEditProtocol):
    def __call__(self, model: LinkInline , context: Context):
        path, _ = model.c[-1]
        start_index = context.text_range.total_count
        for inline in model.inlines:
            context.apply_inline(inline)
        end_index = context.text_range.total_count
        length = end_index - start_index
        if length:
            target = context.text_range.get_range_from_root(start_index + 1, length)
            hyperlink = target.api.ActionSettings(constants.ppMouseClick)
            hyperlink.Action = constants.ppActionHyperlink
            hyperlink.Hyperlink.Address = path
