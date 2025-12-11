from typing import Protocol, runtime_checkable, Callable
from win32com.client import CDispatch


@runtime_checkable
class PPTXObjectProtocol(Protocol):
    @property
    def api(self) -> CDispatch:
        ...

type COMObject = Any
type ObjectLike = COMObject | PPTXObjectProtocol




