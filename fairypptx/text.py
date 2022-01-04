from collections.abc import Sequence
from collections import defaultdict
from pywintypes import com_error
from fairypptx import Shape, Application
from fairypptx import object_utils
from fairypptx.object_utils import is_object, upstream
from fairypptx import registry_utils
from fairypptx.application import Application
from fairypptx import constants
from fairypptx._text import Text, Font, ParagraphFormat
from fairypptx._text.textrange_stylist import ParagraphTextRangeStylist

class TextFrame:
    def __init__(self, arg, *, app=None):
        if app is None:
            self.app = Application()
        else:
            self.app = app

        self._api = self._fetch_api(arg)

    def _fetch_api(self, arg):
        if is_object(arg, "TextFrame"):
            return arg
        elif is_object(arg, "TextFrame2"):
            return arg
        elif isinstance(arg, TextFrame):
            return arg.api
        elif is_object(arg, "Shape"):
            return arg.TextFrame
        elif isinstance(arg, Shape):
            return arg.api.TextFrame
        elif arg is None:
            shape = Shape()
            return shape.api.TextFrame()
        raise ValueError(f"Cannot interpret `arg`; {arg}.")

    @property
    def api(self):
        if is_object(self._api, "TextFrame"):
            return self._api
        elif is_object(self._api, "TextFrame2"):
            return self._to_api(self._api)
        raise RuntimeError("Bug.")

    @property
    def api2(self):
        if is_object(self._api, "TextFrame2"):
            return self._api
        elif is_object(self._api, "TextFrame"):
            return self._to_api2(self._api)
        raise RuntimeError("Bug.")

    def __getattr__(self, name): 
        if "_api" not in self.__dict__:
            raise AttributeError
        return getattr(self.__dict__["_api"], name)

    def __setattr__(self, name, value):
        if "_api" not in self.__dict__:
            object.__setattr__(self, name, value)

        if name in self.__dict__ or name in type(self).__dict__:
            object.__setattr__(self, name, value)
        elif hasattr(self.api, name):
            setattr(self.api, name, value)
        else:
            # TODO: Maybe require modification. 
            object.__setattr__(self, name, value)

    def _to_api2(self, api):
        if is_object(api, "TextFrame2"):
            return api
        elif is_object(api, "TextFrame"):
            return api.Parent.TextFrame2
        raise ValueError()

    def _to_api(self, api):
        if is_object(api, "TextFrame"):
            return api
        elif is_object(api, "TextFrame2"):
            return api.Parent.TextFrame2
        raise ValueError()

"""
[TODO] which is better? 
At `TextRange`  class, internally, TextRange Objects are stored  by `sequence`. 
"""

class TextRange:
    def __init__(self, arg=None, *, app=None):
        if app is None:
            self.app = Application()
        else:
            self.app = app

        self._api = self._fetch_api(arg)

    def _fetch_api(self, arg):
        if is_object(arg, "TextRange"):
            return arg
        elif is_object(arg, "TextRange2"):
            return arg
        elif is_object(arg, "Shape"):
            return arg.TextFrame.TextRange
        elif isinstance(arg, Shape):
            return arg.api.TextFrame.TextRange
        elif arg is None:
            App = self.app.api
            try:
                Selection = App.ActiveWindow.Selection
            except com_error as e: 
                # May be `ActiveWindow` does not exist. (esp at an empty file.)
                pass
            else:
                if Selection.Type == constants.ppSelectionShapes:
                    shape_objects = [shape for shape in Selection.ShapeRange]
                    if len(shape_objects) == 1:
                        return shape_objects[0].TextFrame.TextRange
                    else:
                        raise ValueError("TextRange cannot be generated by multiple shapes.")
                elif Selection.Type == constants.ppSelectionText:
                    return Selection.TextRange
            raise ValueError("Cannot infer an appropriate TextRange.")
        raise ValueError(f"Cannot interpret `arg`; {arg}.")

    @property
    def api(self):
        if is_object(self._api, "TextRange"):
            return self._api
        elif is_object(self._api, "TextRange2"):
            try:
                start = self._api.Start
                length = self._api.Length
                shape_api = upstream(self._api, "Shape")
                return shape_api.TextFrame.TextRange.Characters(start, length)
            except Exception as e:
                raise ValueError from e
        raise RuntimeError("Bug.")

    @property
    def api2(self):
        if is_object(self._api, "TextRange2"):
            return self._api
        elif is_object(self._api, "TextRange"):
            start = self._api.Start
            length = self._api.Length
            shape_api = upstream(self._api, "Shape")
            return shape_api.TextFrame2.TextRange.GetCharacters(start, length)
        raise RuntimeError("Bug.")


    @property
    def shape(self):
        return Shape(upstream(self.api, "Shape"))

    @property
    def characters(self):
        return [TextRange(elem) for elem in self.api.Characters()]

    @property
    def words(self):
        return [TextRange(elem) for elem in self.api.Words()]

    @property
    def lines(self):
        return [TextRange(elem) for elem in self.api.Lines()]

    @property
    def sentences(self):
        return [TextRange(elem) for elem in self.api.Sentences()]

    @property
    def paragraphs(self):
        return [TextRange(elem) for elem in self.api.Paragraphs()]

    @property
    def runs(self):
        return [TextRange(elem) for elem in self.api.Runs()]

    def insert(self, text, mode="after"):
        """Insert the text.
        [TODO] Survey the specification.
        """
        mode = mode.lower()
        assert mode in {"after", "before"}
        insert_funcs = dict()
        insert_funcs["after"] = self.api.InsertAfter
        insert_funcs["before"] = self.api.InsertBefore
        insert_func = insert_funcs[mode]

        s = str(text)
        api_object = insert_func(s)
        tr = TextRange(api_object)
        tr.text  = text
        return tr

    @property
    def text(self):
        return Text(self)

    @text.setter
    def text(self, arg):
        text = Text(arg)
        self.api.Text = str(text)
        self.font = text.font
        self.paragraphformat = text.paragraphformat

    @property
    def font(self):
        return Font(self.api.Font)

    @font.setter
    def font(self, param):
        for key, value in param.items():
            object_utils.setattr(self.api.Font, key, value)

    @property
    def paragraphformat(self):
        return ParagraphFormat(self.api.ParagraphFormat)

    @paragraphformat.setter
    def paragraphformat(self, param):
        for key, value in param.items():
            object_utils.setattr(self.api.ParagraphFormat, key, value)

    
    def itemize(self):
        for elem in self.paragraphs:
            elem.api.ParagraphFormat.Bullet.Visible = constants.msoTrue
            elem.api.ParagraphFormat.Bullet.Type = constants.ppBulletUnnumbered


    def register(self, key, disk=True):
        """ Currently, depending of Paragraphs,
        Style specification rule is ambiguous. 
        Here, (IndentLevel, #paragraphs)'s format is stored.
        Well, then, I wonder whether other mode is introduced or not.

        """
        formatter = ParagraphTextRangeStylist(self)
        registry_utils.register("TextRange", key, formatter, extension=".pkl", disk=disk)

    def like(self, style):
        if isinstance(style, str):
            formatter = registry_utils.fetch("TextRange", style)
            formatter(self)
            return self
        else:
            raise ValueError("Cannot handle, yet.")

    @classmethod
    def make(cls, arg):
        shape = Shape.make(constants.msoShapeRectangle)
        shape.textrange = arg
        return shape.textrange

    @classmethod
    def make_itemization(cls, arg, format=None):
        assert format is None, "Current Implementation"

        """ [TODO]: I'd like a (crude) markdown conversion?
        """
        assert isinstance(arg, Sequence), "Current Implemenation"
        assert all(isinstance(elem, str) for elem in arg), "Current Implementation"
        shape = Shape.make(constants.msoShapeRectangle)
        shape.api.TextFrame.TextRange.Text = "\r".join(arg)
        tr = TextRange(shape)
        tr.api.ParagraphFormat.Bullet.Visible = True
        tr.api.ParagraphFormat.Bullet.Type = constants.ppBulletUnnumbered

        # Itemization's normal display.
        tr.api.ParagraphFormat.Alignment = constants.ppAlignLeft
        return tr


    def __getattr__(self, name): 
        if "_api" not in self.__dict__:
            raise AttributeError
        try:
            return getattr(self.__dict__["_api"], name)
        except AttributeError:
            pass
        # Font's direct access is also possible for this class.
        return getattr(self.__dict__["_api"].Font, name)

    def __setattr__(self, name, value):
        if "_api" not in self.__dict__:
            object.__setattr__(self, name, value)

        if name in self.__dict__ or name in type(self).__dict__:
            object.__setattr__(self, name, value)
        elif hasattr(self.api, name):
            setattr(self.api, name, value)
        elif hasattr(self.api.Font, name):
            # Font's direct access is also possible for this class.
            setattr(self.api.Font, name, value)
        else:
            # TODO: Maybe require modification. 
            object.__setattr__(self, name, value)


# To solve the priority of importing.
from fairypptx._text.editor import MarginMaintainer 
