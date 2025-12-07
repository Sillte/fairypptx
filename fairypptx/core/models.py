from fairypptx.core.types import COMObject


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