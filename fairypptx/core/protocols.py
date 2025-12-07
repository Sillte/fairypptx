from typing import Protocol, runtime_checkable
from win32com.client import CDispatch


@runtime_checkable
class PPTXObjectProtocol(Protocol):
    @property
    def api(self) -> CDispatch:
        ...

