from fairypptx.core.types import COMObject
from typing import Sequence, Literal
from fairypptx.object_utils import stored, getattr as f_getattr, setattr as f_setattr
from fairypptx import registry_utils
from fairypptx import constants
from PIL import Image

def swap_props(api1: COMObject, api2: COMObject, attrs: Sequence[str]) -> None:
    attrs = ["Left", "Top"]
    ps1 = [f_getattr(api1, attr) for attr in attrs]
    ps2 = [f_getattr(api2, attr) for attr in attrs]
    for attr, p1, p2 in zip(attrs, ps1, ps2):
        f_setattr(api1, attr, p2)
        f_setattr(api2, attr, p1)


def tighten(api: COMObject, *, oneline: bool=False):
    """Tighten the Shape according to Text.

    Args:
        oneline: Modify so that text becomes 1 line.
    """
    if api.HasTextFrame:
        if oneline is True:
            api.TextFrame.TextRange.Text = api.Text.replace("\r", "").replace(
                "\n", ""
            )
        with stored(api, ("TextFrame.AutoSize", "TextFrame.WordWrap")):
            api.TextFrame.AutoSize = constants.ppAutoSizeShapeToFitText
            api.TextFrame.WordWrap = constants.msoFalse

def is_tight(api: COMObject, *, oneline: bool=False):
    assert oneline is False
    width, height = api.Width, api.Height
    with stored(api, ("Width", "Height", "Left", "Top")):
        tighten(api, oneline=False)
        if abs(width - api.Width) <= 5 and abs(height - api.Height) <= 5:
            return True
    return False

def to_image(api:COMObject, mode: Literal["RGBA", "RGB"] ="RGBA") -> Image.Image:
    with registry_utils.yield_temporary_path(suffix=".png") as path:
        api.Export(path, constants.ppShapeFormatPNG)
        image = Image.open(path).copy()
    return image.convert(mode)
