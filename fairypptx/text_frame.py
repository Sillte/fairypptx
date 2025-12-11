import itertools
from collections.abc import Sequence
from collections import defaultdict
from pywintypes import com_error
from fairypptx import Shape, Application
from fairypptx import object_utils
from fairypptx.object_utils import is_object, upstream
from fairypptx import registry_utils
from fairypptx.core.application import Application
from fairypptx import constants
from fairypptx.font import Font
from fairypptx.paragraph_format import ParagraphFormat

from fairypptx.core.resolvers import resolve_textframe
from fairypptx.core.types import COMObject, PPTXObjectProtocol  

class TextFrame:
    def __init__(self, arg):
        self._api = resolve_textframe(arg)

    @property
    def api(self) -> COMObject:
        return self._api


    @property
    def api2(self) -> COMObject:
        return self._api.Parent.TextFrame2

