from pathlib import Path
from typing import Sequence, Any
from fairypptx.core.protocols import PPTXObjectProtocol
from fairypptx.core.application import Application
from fairypptx.core.types import COMObject
from fairypptx import constants
from fairypptx.object_utils import is_object
from collections import UserString

from win32com.client import DispatchEx, GetActiveObject
from pywintypes import com_error

def get_application_api() -> COMObject:
    """Return Application API.
    """
    try:
        api = GetActiveObject("Powerpoint.Application")
    except com_error:
        api = DispatchEx("Powerpoint.Application")
    return api


def to_api_or_none(arg: Any) -> None | COMObject:
    """Return `COMOBject`, if possible. 
    Otherwise return None. 
    """
    if isinstance(arg, PPTXObjectProtocol):
        return arg.api
    if is_object(arg):
        return arg
    return None


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

def resolve_shape_range(arg: PPTXObjectProtocol | COMObject | None = None) -> COMObject:
    if isinstance(arg, PPTXObjectProtocol) or is_object(arg):
        if isinstance(arg, PPTXObjectProtocol):
            api: COMObject = arg.api 
        else:
            api = arg
        if is_object(api, "ShapeRange"):
            return api

    if arg is None:
        App = Application().api
        try:
            Selection = App.ActiveWindow.Selection
        except com_error:
            # May be `ActiveWindow` does not exist. (esp at an empty file.)
            pass
        else:
            if Selection.Type == constants.ppSelectionShapes:
                if not Selection.HasChildShapeRange:
                    return Selection.ShapeRange
                else:
                    return Selection.ChildShapeRange
            elif Selection.Type == constants.ppSelectionText:
                # Even if Seleciton.Type is ppSelectionText, `Selection.ShapeRange` return ``Shape``.
                return Selection.ShapeRange
        shapes_api = resolve_shapes()
        return shapes_api.Range()
    raise NotImplementedError()


def resolve_shapes(arg: PPTXObjectProtocol | COMObject | None = None) -> COMObject:
    if isinstance(arg, PPTXObjectProtocol) or is_object(arg):
        if isinstance(arg, PPTXObjectProtocol):
            api: COMObject = arg.api 
        else:
            api = arg

        if is_object(api, "ShapeRange"):
            return api.Parent.Shapes
        elif is_object(api, "Shapes"):
            return api
        elif is_object(api, "Slide"):
            return api.Shapes
        elif is_object(api, "Shape"):
            return api.Parent.Shapes

    if isinstance(arg, Sequence):
        apis = [to_api_or_none(elem) for elem in arg]
        if not apis[0]:
            raise ValueError(f"Cannot interpret `arg`; {arg}.") 
        if apis[0] is None:
            raise ValueError(f"Cannot interpret `arg`; {arg}.") 
        
        # [TODO] This judge may be too loose....
        return apis[0].Parent.Shapes

    if arg is None:
        App = Application().api
        try:
            Selection = App.ActiveWindow.Selection
        except com_error as e:
            pass
        else:
            if Selection.Type == constants.ppSelectionShapes:
                if Selection.HasChildShapeRange:
                    shape_objects = [shape for shape in Selection.ChildShapeRange]
                else:
                    shape_objects = [shape for shape in Selection.ShapeRange]
                return shape_objects[0].Parent.Shapes
    return resolve_slide().Shapes


def resolve_shape(arg: PPTXObjectProtocol | COMObject | None = None) -> COMObject:
    if isinstance(arg, PPTXObjectProtocol) or is_object(arg):
        if isinstance(arg, PPTXObjectProtocol):
            api: COMObject = arg.api 
        else:
            api = arg
        if is_object(api, "Shape"):
            return api


    if isinstance(arg, Sequence):
        apis = [to_api_or_none(elem) for elem in arg]
        apis = [elem for elem in apis if elem is not None]
        if not apis[0]:
            raise ValueError(f"Cannot interpret `arg`; {arg}.") 
        if apis[0] is None:
            raise ValueError(f"Cannot interpret `arg`; {arg}.") 
        
        # [TODO] This judge may be too loose....
        return apis[0]

    if arg is None:
        App = Application().api
        try:
            Selection = App.ActiveWindow.Selection
        except com_error as e:
            pass
        else:
            if Selection.Type == constants.ppSelectionShapes:
                if Selection.HasChildShapeRange:
                    shape_objects = [shape for shape in Selection.ChildShapeRange]
                else:
                    shape_objects = [shape for shape in Selection.ShapeRange]
                if shape_objects:
                    return shape_objects[0]
    shapes_objects = resolve_slide().Shapes
    if shapes_objects.Count:
        return shapes_objects.Item(1)
    raise ValueError("Cannot obtain `Shape` api.")


def resolve_table(arg: PPTXObjectProtocol | COMObject | None = None) -> COMObject:
    if isinstance(arg, PPTXObjectProtocol) or is_object(arg):
        if isinstance(arg, PPTXObjectProtocol):
            api: COMObject = arg.api 
        else:
            api = arg
        if is_object(api, "Table"):
            return api
        if is_object(api, "Shape"):
            return api.Table

    if arg is None:
        shape_api = resolve_shape(None)
        return shape_api.Table

    raise ValueError(f"Cannot interpret `arg`; {arg}.")



def resolve_textframe(arg: PPTXObjectProtocol | COMObject | None = None) -> COMObject:
    if isinstance(arg, PPTXObjectProtocol) or is_object(arg):
        if isinstance(arg, PPTXObjectProtocol):
            api: COMObject = arg.api 
        else:
            api = arg
        if is_object(api, "TextFrame"):
            return api
        if is_object(api, "TextFrame2"):
            return api.Parent.TextFrame
        if is_object(api, "Shape"):
            return api.TextFrame
    shape_api = resolve_shape(None)
    return shape_api.TextFrame

def resolve_text_range(arg: PPTXObjectProtocol | COMObject | None = None) -> COMObject:
    if isinstance(arg, PPTXObjectProtocol) or is_object(arg):
        if isinstance(arg, PPTXObjectProtocol):
            api: COMObject = arg.api 
        else:
            api = arg
        if is_object(api, "TextRange"):
            return api
        if is_object(api, "TextRange2"):
            return api.Parent.TextRange
        if is_object(api, "Shape"):
            return api.TextFrame.TextRange

    if arg is None:
        App = Application().api
        try:
            Selection = App.ActiveWindow.Selection
        except com_error as e: 
            # May be `ActiveWindow` does not exist. (esp at an empty file.)
            pass
        else:
            if Selection.Type == constants.ppSelectionShapes:
                shape_objects = [shape_obj for shape_obj in Selection.ShapeRange]
                if len(shape_objects) == 1:
                    return shape_objects[0].TextFrame.TextRange
                else:
                    raise ValueError("TextRange cannot be generated by multiple shapes.")
            elif Selection.Type == constants.ppSelectionText:
                # When Selection is 0, then
                # It is considered that entire `TextRange` is selected. 
                textrange_api = Selection.TextRange
                if 0 < textrange_api.Length:
                    return textrange_api
                else:
                    return textrange_api.Parent.TextRange
    raise ValueError(f"Cannot interpret `arg`; {arg}.")


          

