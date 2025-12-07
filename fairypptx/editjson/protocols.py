from typing import Self, runtime_checkable
from typing import Protocol
from fairypptx.core.types import PPTXObjectProtocol


@runtime_checkable
class EditParamProtocol[T: PPTXObjectProtocol](Protocol):
    @classmethod
    def from_entity(cls, entity: T) -> Self:
        """Generate itself from the entity of `fairpptx.PPTXObject`
        """
        ...

    def apply(self, entity: T) -> T:
        """Apply this edit param to 
        """
        ...
    

if __name__ == "__main__":
    pass
