# from comtypes import client
from win32com.client import DispatchEx, GetActiveObject
from win32com.client import CDispatch
from pywintypes import com_error


class Application:
    def __init__(self):
        try:
            api = GetActiveObject("Powerpoint.Application")
        except com_error:
            api = DispatchEx("Powerpoint.Application")
        self._api = api
        self._api.Visible = True

    @property
    def api(self) -> CDispatch:
        return self._api 

