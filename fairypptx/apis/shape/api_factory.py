"""Low-level factory for creating COM shape objects on slides.

This module provides API-level shape creation by obtaining a slide's Shapes
collection and using COM methods to add shapes, pictures, etc.

The factory accepts optional `shapes_api` argument; if omitted, it uses
`resolve_shapes()` to obtain the default Shapes collection.
"""

from typing import Literal
from PIL import Image
from fairypptx.core.resolvers import resolve_shapes
from fairypptx.core.types import COMObject
from fairypptx import constants
from fairypptx.constants import msoTrue, msoFalse
from fairypptx import registry_utils


class ShapeApiFactory:
    """Factory for low-level COM shape object creation.

    Each method takes an optional `shapes_api` parameter. If provided, shapes
    are added to that shapes collection; otherwise `resolve_shapes()` is used
    to obtain the default Shapes collection. This allows callers to control
    which slide receives the shapes while providing a sensible default.
    """

    @staticmethod
    def add_shape_from_type(type_: int, shapes_api: COMObject | None = None, **kwargs) -> COMObject:
        """Create a shape of a given type.

        Args:
            type_: COM shape type constant (e.g., constants.msoShapeRectangle).
            shapes_api: Optional Shapes collection. If None, uses resolve_shapes().
            **kwargs: Additional arguments passed to AddShape (Left, Top, Width, Height, etc.).
                      If not provided, defaults: Left=0, Top=0, Width=100, Height=100.

        Returns:
            The COM shape object.
        """
        if shapes_api is None:
            shapes_api = resolve_shapes() 
        assert shapes_api is not None
        # Set defaults for required COM parameters
        left = kwargs.pop("Left", 0)
        top = kwargs.pop("Top", 0)
        width = kwargs.pop("Width", 100)
        height = kwargs.pop("Height", 100)
        return shapes_api.AddShape(type_, left, top, width, height, **kwargs)

    @staticmethod
    def add_picture(
        image_or_path: Image.Image | str,
        shapes_api: COMObject | None = None,
        **kwargs
    ) -> COMObject:
        """Add a picture to a slide.

        Args:
            image_or_path: PIL.Image or file path string.
            shapes_api: Optional Shapes collection. If None, uses resolve_shapes().
            **kwargs: Additional arguments passed to AddPicture.

        Returns:
            The COM picture shape object.
        """
        if shapes_api is None:
            shapes_api = resolve_shapes() 
        assert shapes_api is not None

        if isinstance(image_or_path, Image.Image):
            with registry_utils.yield_temporary_dump(image_or_path) as path:
                return shapes_api.AddPicture(
                    str(path),
                    msoFalse,
                    msoTrue,
                    Left=0,
                    Top=0,
                    Width=image_or_path.width,
                    Height=image_or_path.height,
                    **kwargs
                )
        else:
            return shapes_api.AddPicture(str(image_or_path), msoFalse, msoTrue, **kwargs)

    @staticmethod
    def add_textbox(
        text: str,
        shapes_api: COMObject | None = None,
        **kwargs
    ) -> COMObject:
        """Create a textbox with text.

        Args:
            text: Text to insert into the shape.
            shapes_api: Optional Shapes collection. If None, uses resolve_shapes().
            **kwargs: Additional arguments passed to AddShape (Left, Top, Width, Height, etc.).
                      If not provided, defaults: Left=0, Top=0, Width=100, Height=100.

        Returns:
            The COM shape object (textbox).
        """
        if shapes_api is None:
            shapes_api = resolve_shapes()
        # Create the rectangle shape (pass through shapes_api and kwargs so
        # position/size can be controlled by the caller)
        shape = ShapeApiFactory.add_shape_from_type(
            constants.msoShapeRectangle, shapes_api=shapes_api, **kwargs
        )
        # `shape` is a COM object; set its text directly and return it.
        shape.TextFrame.TextRange.Text = text
        return shape

    @staticmethod
    def make_arrow(
        direction: Literal["right", "left", "up", "down", "both"] = "right",
        shapes_api: COMObject | None = None,
        **kwargs
    ) -> COMObject:
        """Create an arrow shape in the specified direction.

        Args:
            direction: "right", "left", "up", "down", or "both".
            shapes_api: Optional Shapes collection. If None, uses resolve_shapes().
            **kwargs: Additional arguments passed to AddShape (Left, Top, Width, Height, etc.).

        Returns:
            The COM arrow shape object.
        """
        type_map = {
            "right": constants.msoShapeRightArrow,
            "left": constants.msoShapeLeftArrow,
            "up": constants.msoShapeUpArrow,
            "down": constants.msoShapeDownArrow,
            "both": constants.msoShapeLeftRightArrow,
        }
        type_id = type_map[direction]
        return ShapeApiFactory.add_shape_from_type(type_id, shapes_api, **kwargs)
