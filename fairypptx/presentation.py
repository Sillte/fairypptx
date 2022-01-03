from fairypptx.application import Application
from pathlib import Path
from collections import UserString
from pywintypes import com_error
from fairypptx.object_utils import is_object

class Presentation:
    def __init__(self, arg=None, *, app=None):
        if app is None:
            self.app = Application()
        else:
            self.app = app

        self._api = self._fetch_api(arg)

    @property
    def api(self):
        return self._api

    @property
    def slides(self):
        from fairypptx.slide import Slides
        return Slides(self.api.Slides)

    def _fetch_api(self, arg):
        if is_object(arg, "Presentation"):
            return arg
        elif isinstance(arg, Presentation):
            return arg.api
        elif isinstance(arg, (str, UserString, Path)):
            # print("arg", arg)
            App = self.app.api
            # Check the specified presentation is opened
            arg = Path(arg).absolute()
            path_to_pres = {Path(pres.FullName): pres for pres in App.Presentations}
            if arg in path_to_pres:
                return path_to_pres[arg]
            assert arg.suffix in {".pptm", ".pptx"}
            return App.Presentations.Open(str(arg))

        elif arg is None:
            App = self.app.api
            try:
                return App.ActivePresentation
            except com_error:
                pass

            # Return the first Presentation.
            if App.Presentations.Count:
                return App.Presentations[1]

            # Last resort; add and return.
            return App.Presentations.Add()
        raise ValueError(f"Cannot interpret `arg`; {arg}.")
