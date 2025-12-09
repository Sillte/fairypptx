from fairypptx.editjson.protocols import EditParamProtocol
from fairypptx.editjson.shape import NaiveShapeStyle
from fairypptx.editjson.text_range import NaiveTextRangeParagraphStyle


class PPTXObjectStyleTypeRegistry:
    def __init__(self, default_type: type[EditParamProtocol]) -> None:
        self.registry: dict[str, type[EditParamProtocol]] = {}
        self.default_type = default_type

    def register(
        self,
        edit_param_cls: type[EditParamProtocol],
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

    def fetch(self, cls_name: str | None = None) -> type[EditParamProtocol]:
        if cls_name is None:
            return self.default_type
        if cls_name not in self.registry:
            msg = f"Unknown style type: {cls_name!r}"
            raise KeyError(msg)
        return self.registry[cls_name]


ShapeStyleTypeRegistry = PPTXObjectStyleTypeRegistry(NaiveShapeStyle)
TextRangeStyleTypeRegistry = PPTXObjectStyleTypeRegistry(NaiveTextRangeParagraphStyle)

