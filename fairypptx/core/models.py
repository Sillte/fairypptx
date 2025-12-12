from fairypptx.core.types import COMObject
from fairypptx.core.protocols import PPTXObjectProtocol
from typing import Any, Callable 

from fairypptx.object_utils import is_object


from pydantic import BaseModel


from abc import ABC, abstractmethod
from typing import Self


class BaseApiModel(BaseModel, ABC):
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
        

class ApiApplicator[T]:
    def __init__(self,  api_model_type: type[BaseApiModel], apply_custom: Callable[[COMObject, T], None] | None = None):
        self._api_model_type = api_model_type
        self._apply_custom = apply_custom

    def apply_api(self, api: COMObject, value: COMObject) -> None:
        api_bridge = self._api_model_type.from_api(value)
        api_bridge.apply_api(api)

    def apply(self, api: COMObject, value: T) -> None:
        if isinstance(value, PPTXObjectProtocol):
            self.apply_api(api, value.api)
        elif is_object(value):
            self.apply_api(api, value)
        else: 
            if self._apply_custom:
                self._apply_custom(api, value)
            else:
                msg = f"`{value=}` cannot be handled for `{type(self)}`"
                raise ValueError(msg)
