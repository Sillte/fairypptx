from fairypptx.styles.protocols import StyleModelProtocol
from fairypptx.styles.shape import NaiveShapeStyle
from fairypptx.styles.text_range import NaiveTextRangeParagraphStyle
from fairypptx.styles.table import NaiveTableStyle


class PPTXObjectStyleTypeRegistry:
    def __init__(self, default_type: type[StyleModelProtocol]) -> None:
        self.registry: dict[str, type[StyleModelProtocol]] = {}
        self.default_type = default_type

    def register(
        self,
        edit_param_cls: type[StyleModelProtocol],
        cls_name: str | None = None,
        *,
        override_default: bool = False
    ) -> None:

        if cls_name is None:
            cls_name = edit_param_cls.__name__
        assert isinstance(cls_name, str)

        self.registry[cls_name] = edit_param_cls

        if override_default:
            self.default_type = edit_param_cls

    def fetch(self, cls_name: str | None = None) -> type[StyleModelProtocol]:
        if cls_name is None:
            return self.default_type
        if cls_name not in self.registry:
            msg = f"Unknown style type: {cls_name!r}"
            raise KeyError(msg)
        return self.registry[cls_name]


ShapeStyleTypeRegistry = PPTXObjectStyleTypeRegistry(NaiveShapeStyle)
TextRangeStyleTypeRegistry = PPTXObjectStyleTypeRegistry(NaiveTextRangeParagraphStyle)
TableStyleTypeRegistry = PPTXObjectStyleTypeRegistry(NaiveTableStyle)

