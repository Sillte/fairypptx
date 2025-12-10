from fairypptx.core.types import COMObject
from typing import Any 


from pydantic import BaseModel


from abc import ABC, abstractmethod
from typing import Self


class ApiBridgeBaseModel(BaseModel, ABC):
    @classmethod
    @abstractmethod
    def from_api(cls, api: COMObject) -> Self:
        ...

    @abstractmethod
    def apply_api(self, api: COMObject) -> COMObject:
        ...

    def __eq__(self, other: Any) -> bool:
        if not isinstance(other, self.__class__):
            return False
        return (
            self.model_dump(exclude_defaults=True)
            == other.model_dump(exclude_defaults=True)
        )