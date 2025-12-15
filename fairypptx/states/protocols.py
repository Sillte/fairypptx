from typing import Self, runtime_checkable
from typing import Protocol
from win32com.client import CDispatch
from fairypptx.core.types import PPTXObjectProtocol


@runtime_checkable
class PPTXEntityProtocol(Protocol):
    @property
    def id(self) -> int:
        ...

    @property
    def api(self) -> CDispatch:
        ...
 

    

if __name__ == "__main__":
    pass
