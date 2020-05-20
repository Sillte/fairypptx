from comtypes import client


class Application:
    def __init__(self):
        self.api = client.CreateObject("Powerpoint.Application")
        self.api.Visible = True

    def __getattr__(self, name): 
        return getattr(self.api, name)

    """
    @property
    def presentation(self):
        try:
            return Presentation(self.api.ActivePresentation)
        except COMError:
            pass

        # Return the first Presentation.
        if self.api.Presentations.Count:
            return self.api.Presentations[1]

        # Last resort; add and return.
        pres = self.api.Presentations.Add()
        return Presentation(pres)
    """
        

