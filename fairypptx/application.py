# from comtypes import client
from win32com.client import DispatchEx, GetActiveObject
from pywintypes import com_error


class Application:
    def __init__(self):
        try:
            api = GetActiveObject("Powerpoint.Application")
        except com_error:
            api = DispatchEx("Powerpoint.Application")
        self.api = api
        self.api.Visible = True

    def __getattr__(self, name): 
        return getattr(self.api, name)

    """
    @property
    def presentation(self):
        try:
            return Presentation(self.api.ActivePresentation)
        except com_error:
            pass

        # Return the first Presentation.
        if self.api.Presentations.Count:
            return self.api.Presentations[1]

        # Last resort; add and return.
        pres = self.api.Presentations.Add()
        return Presentation(pres)
    """
        

