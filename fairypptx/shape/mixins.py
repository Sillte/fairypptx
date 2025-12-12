from fairypptx.core.types import PPTXObjectProtocol


class LocationMixin:
    """This Mixin handles the functionality of geometry information of `Shape`.
    This Mixin must be applicable to all the `Shape` in the domain of COMObject.
    """

    @property
    def left(self: PPTXObjectProtocol) -> float:
        return self.api.Left

    @left.setter
    def left(self: PPTXObjectProtocol, value: float) -> None:
        self.api.Left = value

    @property
    def top(self: PPTXObjectProtocol) -> float:
        return self.api.Top

    @top.setter
    def top(self: PPTXObjectProtocol, value: float) -> None:
        self.api.Top = value

    @property
    def width(self: PPTXObjectProtocol) -> float:
        return self.api.Width

    @width.setter
    def width(self: PPTXObjectProtocol, value: float) -> None:
        self.api.Width = value

    @property
    def height(self: PPTXObjectProtocol) -> float:
        return self.api.Height

    @height.setter
    def height(self: PPTXObjectProtocol, value: float) -> None:
        self.api.Height = value

    @property
    def size(self: PPTXObjectProtocol) -> tuple[float, float]:
        return (self.api.Width, self.api.Height)

    @size.setter
    def size(self: PPTXObjectProtocol, value: tuple[float, float]) -> None:
        self.api.Width, self.api.Height = value

    @property
    def rotation(self: PPTXObjectProtocol) -> float:
        return self.api.Rotation

    @rotation.setter
    def rotation(self: PPTXObjectProtocol, value: float) -> None:
        self.api.Rotation = value

    def rotate(self: PPTXObjectProtocol, degree: float) -> None:
        self.api.Rotation += degree
