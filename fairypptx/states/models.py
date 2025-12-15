from fairypptx.states.protocols import PPTXEntityProtocol

from typing import Any,  Annotated

from pydantic import BaseModel, Field


from abc import ABC, abstractmethod
from typing import Self


class BaseStateModel[T: PPTXEntityProtocol](BaseModel, ABC):
    id: Annotated[int, Field(description="Indentifier of the entity.")]

    @classmethod
    @abstractmethod
    def from_entity(cls, entity: T) -> Self:
        ...

    @abstractmethod
    def apply(self, entity: T):
        ...

    def __eq__(self, other: Any) -> bool:
        if not isinstance(other, self.__class__):
            return False
        return (
            self.model_dump(exclude_defaults=True)
            == other.model_dump(exclude_defaults=True)
        )
