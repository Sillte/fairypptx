from fairypptx.apis.paragraph_format.api_model import ParagraphFormatApiModel
from fairypptx.apis.paragraph_format.applicator import ParagraphFormatApplicator
from fairypptx.core.types import COMObject
from fairypptx.object_utils import to_api2


from fairypptx.object_utils import to_api2, is_object


class ParagraphFormat:
    """Represents the Paragraph Information. """

    def __init__(self, api: COMObject) -> None:
        if isinstance(api, ParagraphFormat):
            api = api.api
        assert is_object(api)
        self._api = api
        
    @property
    def api(self) -> COMObject:
        return self._api

    @property
    def api2(self) -> COMObject:
        return to_api2(self._api)

    @property
    def indent_level(self):
        return self.api2.IndentLevel

    @indent_level.setter
    def indent_level(self, value: int):
        self.api2.IndentLevel = value

    @property
    def alignment(self) -> int:
        return self.api2.Alignment

    @alignment.setter
    def alignment(self, value: int):
        self.api2.Alignment = value

    @property
    def space_before(self) -> float:
        return self.api2.SpaceBefore

    @space_before.setter
    def space_before(self, value: float):
        self.api2.SpaceBefore = value

    @property
    def bullet_type(self) -> int | None:
        if self.api2.Bullet.Visible:
            return self.api2.Bullet.Type
        return None

    @bullet_type.setter
    def bullet_type(self,  value: int | None):
        if value is None:
            self.api2.Bullet.Visible = False
        else:
            self.api2.Bullet.Visible = True
            self.api2.Bullet.Type = value
    

class ParagraphFormatProperty:
    def __get__(self, parent: COMObject, objtype=None):
        return ParagraphFormat(parent.api.ParagraphFormat)


    def __set__(self, shape, value):
        ParagraphFormatApplicator.apply(shape.api.ParagraphFormat, value)



if __name__ == "__main__":
    pass
