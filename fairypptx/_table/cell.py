from fairypptx.core.types import COMObject

class Cell:
    def __init__(self, api:COMObject) -> None:
        self._api = api
        
    @property
    def api(self) -> COMObject:
        return self._api

    @property
    def shape(self) -> "Shape":
        from fairypptx.shape import Shape
        return Shape(self.api.Shape)

    @property
    def text(self):
        return self.shape.text

    def is_empty(self):
        """ 
        (2020-04-25): Currently, only text is handled.
        """
        return (self.shape.text == "")