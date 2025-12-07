from typing import Self, runtime_checkable
from typing import Protocol
from fairypptx.core.types import PPTXObjectProtocol, COMObject
from pydantic import BaseModel
from abc import ABC, abstractmethod


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
    

class ApiApplyBaseModel(BaseModel, ABC):
    @classmethod
    @abstractmethod
    def from_api(cls, api: COMObject) -> Self:
        ...

    @abstractmethod
    def apply_api(self, api: COMObject) -> COMObject:
        ...


if __name__ == "__main__":
    pass
