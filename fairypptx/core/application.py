from win32com.client import DispatchEx, GetActiveObject
from fairypptx.core.types import COMObject  
from pywintypes import com_error


class Application:
    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialize()
        return cls._instance

    def _initialize(self):
        try:
            api = GetActiveObject("Powerpoint.Application")
        except com_error:
            api = DispatchEx("Powerpoint.Application")

        self._api = api
        self._api.Visible = True

    @property
    def api(self) -> COMObject:
        return self._api
