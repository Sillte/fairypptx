from pathlib import Path
from fairypptx.core.protocols import PPTXObjectProtocol
from fairypptx.core.application import Application
from fairypptx.core.types import COMObject
from fairypptx import constants
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


def resolve_slide(arg: PPTXObjectProtocol | COMObject | None = None) -> COMObject:
    """Return the COMObject of `Slide`."""

    if isinstance(arg, PPTXObjectProtocol) or is_object(arg):
        if isinstance(arg, PPTXObjectProtocol):
            api: COMObject = arg.api 
        else:
            api = arg
        if is_object(api, "Slide"):
            return api
        msg = f"`{arg}` is not acceptable for `Slide`."
        raise ValueError(msg)

    if arg is None:
        App = Application().api
        try:
            if App.ActiveWindow.ViewType != constants.ppViewNormal:
                App.ActiveWindow.ViewType = constants.ppViewNormal
        except com_error:
            pass
        try:
            if App.ActiveWindow.Selection.SlideRange:
                return App.ActiveWindow.Selection.SlideRange.Item(1)
        except com_error:
            pass

        Pres = resolve_presentation(None)
        if Pres.Slides.Count:
            return Pres.Slides(Pres.Slides.Count)
        else:
            return Pres.Slides.Add(1, constants.ppLayoutBlank)

def resolve_slide_range(arg: PPTXObjectProtocol | COMObject | None = None) -> COMObject:
    """Return the COMObject of `Slide`."""

    if isinstance(arg, PPTXObjectProtocol) or is_object(arg):
        if isinstance(arg, PPTXObjectProtocol):
            api: COMObject = arg.api 
        else:
            api = arg
        if is_object(api, "SlideRange"):
            return api
        msg = f"`{arg}` is not acceptable for `SlideRange`."
        raise ValueError(msg)
    if arg is None:
        App = Application().api
        if App.ActiveWindow.Selection.SlideRange:
            return App.ActiveWindow.Selection.SlideRange
        else:
            return App.ActivePresentation.Slides.Range()

def resolve_slides(arg: PPTXObjectProtocol | COMObject | None = None) -> COMObject:
    """Return the COMObject of `Slides`."""

    if isinstance(arg, PPTXObjectProtocol) or is_object(arg):
        if isinstance(arg, PPTXObjectProtocol):
            api: COMObject = arg.api 
        else:
            api = arg
        if is_object(api, "Slides"):
            return api
        msg = f"`{arg}` is not acceptable for `Slides`."
        raise ValueError(msg)
    if arg is None:
        App = Application().api
        return App.ActivePresentation.Slides
