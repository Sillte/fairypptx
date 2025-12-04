from pathlib import Path
from fairypptx.core.protocols import PPTXObjectProtocol
from fairypptx.core.types import COMObject
from fairypptx.object_utils import is_object
from collections import UserString

from win32com.client import DispatchEx, GetActiveObject
from pywintypes import com_error

def get_application_api() -> COMObject:
    try:
        api = GetActiveObject("Powerpoint.Application")
    except com_error:
        api = DispatchEx("Powerpoint.Application")
    return api


def resolve_presentation(arg: PPTXObjectProtocol | COMObject | None | str | Path | UserString) -> COMObject:
    """Return the COMObject of `Presentation`.
    """
    if isinstance(arg, PPTXObjectProtocol) or is_object(arg):
        if isinstance(arg, PPTXObjectProtocol):
            api: COMObject = arg.api 
        else:
            api = arg
        if is_object(api, "Presentation"):
            return api
        msg = f"`{arg}` is not acceptable for `Presentation`."
        raise ValueError(msg)

    if isinstance(arg, (str, Path, UserString)):
        # print("arg", arg)
        App = get_application_api()
        # Check the specified presentation is opened
        if isinstance(arg, UserString):
            arg = str(arg)
        arg = Path(arg).absolute()
        path_to_pres = {Path(pres.FullName): pres for pres in App.Presentations}
        if arg in path_to_pres:
            return path_to_pres[arg]
        assert arg.suffix in {".pptm", ".pptx"}, "Cannot handle this file."
        return App.Presentations.Open(str(arg))
    elif arg is None:
        App = get_application_api()
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

