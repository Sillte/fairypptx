"""Handling processing related to location and position. 

"""

import numpy as np
from collections import defaultdict
from fairypptx.box import Box
from fairypptx.slide import GridHandler


class ShapesLocator:
    """Locate `Shape` / `Shapes`

    Args:
        mode
        `blank`:
               To the center of maximum black are in the slide.
        `center`:
               To the center of the slide.
    """

    def __init__(self, mode: str = "blank"):
        self.mode = mode.lower()

    def __call__(self, arg):
        shapes = self._to_shapes(arg)
        if self.mode == "center":
            self._to_center(shapes)
        elif self.mode == "blank":
            self._to_blank_area(shapes)
        else:
            raise ValueError(f"Invalid mode `{self.mode}`.")
        return arg 

    def _to_shapes(self, arg):
        """Convert to `Shapes`."""
        from fairypptx.shape import Shape
        from fairypptx.shapes import Shapes
        from fairypptx.object_utils import is_object
        from typing import Sequence

        if isinstance(arg, Shapes):
            return arg
        elif isinstance(arg, Sequence):
            return Shapes(arg)
        elif isinstance(arg, Shape):
            return Shapes([arg])
        elif is_object(arg, "Shapes"):
            return Shapes(arg)
        elif is_object(arg, "Shape"):
            return Shapes(arg)
        raise ValueError(f"Cannot interpret `{arg}`.")

    def _to_blank_area(self, shapes):
        shapes = self._to_shapes(shapes)
        remove_ids = set(shape.api.Id for shape in shapes)
        grid_handler = GridHandler(shapes.slide)
        target_shapes = [
            shape
            for shape in grid_handler.slide.shapes
            if shape.api.Id not in remove_ids
        ]

        r_occupations = grid_handler.make_occupations(target_shapes)
        canvas = grid_handler.get_maximum_box(r_occupations)

        box = shapes.circumscribed_box
        x_margin = max((canvas.width - box.width) / 2, 0)
        y_margin = max((canvas.height - box.height) / 2, 0)

        dx = (canvas.left + x_margin) - box.left
        dy = (canvas.top + y_margin) - box.top
        for shape in shapes:
            shape.left += dx
            shape.top += dy
        return shapes

    def _to_center(self, shapes):
        shapes = self._to_shapes(shapes)
        c_box = shapes.circumscribed_box
        target_width = c_box.width
        target_height = c_box.height
        slide_width = shapes[0].slide.box.width
        slide_height = shapes[0].slide.box.height
        left = (slide_width - target_width) / 2
        top = (slide_height - target_height) / 2
        return self._locate_shapes(shapes, left, top)

    def _move_shapes(self, shapes, dx, dy):
        """Move `Shapes`"""
        raise NotImplementedError("")

    def _locate_shapes(self, shapes, left, top):
        """Locate `Shapes` so that the circumscribed box's
        left and top becomes as specified.
        """
        current_box = shapes.circumscribed_box
        current_left = current_box.left
        current_top = current_box.top
        for shape in shapes:
            shape.left += left - current_left
            shape.top += top - current_top
        return shapes


class ShapesAdjuster:
    """Adjust `Shapes`.

    The interval changes depending of the given situation.
    Specifically, the decision of the circumscribed box differs.

    * `is_edge_interval = True`:
    One `Shape` enclose all the shapes. the circumscribed box is this shapes's.
    * `is_edge_interval = False`:
    The circumscribed box is determined by all the shapes.
    """

    def __init__(self, axis=None):
        self.axis = axis

    def _yield_axis(self, axis, shapes):
        if axis == "width":
            axis = 1
        if axis == "height":
            axis = 0
        if axis is None:
            boxes = [shape.box for shape in shapes]
            center_ys = [box.center[0] for box in boxes]
            center_xs = [box.center[1] for box in boxes]
            if np.std(center_xs) < np.std(center_ys):
                axis = 0
            else:
                axis = 1
        assert axis in {0, 1}
        return axis

    def _yield_circumscribed_box(self, shapes):
        boxes = [shape.box for shape in shapes]
        c_left = min(box.left for box in boxes)
        c_top = min(box.top for box in boxes)
        c_right = max(box.right for box in boxes)
        c_bottom = max(box.bottom for box in boxes)
        c_box = Box(left=c_left, top=c_top, width=c_right - c_left, height=c_bottom - c_top)
        return c_box

    def _adjust_horizontally(self, shapes, c_box, is_edge_interval):

        boxes = [shape.box for shape in shapes]
        # c -> circumscribed
        c_left = c_box.left
        c_right = c_box.right

        r_width = c_right - c_left
        s_width = sum(box.width for box in boxes)
        # `n_interval` and offset setting is
        if not is_edge_interval:
            n_interval = len(shapes) - 1
            interval_width = (r_width - s_width) / n_interval
            current_x = c_left
        else:
            n_interval = len(shapes) + 1
            interval_width = (r_width - s_width) / n_interval
            current_x = c_left + interval_width
        shapes = sorted(shapes, key=lambda shape: shape.left)
        for index, shape in enumerate(shapes):
            shape.left = current_x
            current_x += shape.width + interval_width

    def _adjust_vertially(self, shapes, c_box, is_edge_interval):
        boxes = [shape.box for shape in shapes]

        c_top = c_box.top
        c_bottom = c_box.bottom

        r_height = c_bottom - c_top
        s_height = sum(box.height for box in boxes)

        if not is_edge_interval:
            n_interval = len(shapes) - 1
            interval_height = (r_height - s_height) / n_interval
            current_y = c_top
        else:
            n_interval = len(shapes) + 1
            interval_height = (r_height - s_height) / n_interval
            current_y = c_top + interval_height
        shapes = sorted(shapes, key=lambda shape: shape.top)
        for index, shape in enumerate(shapes):
            shape.top = current_y
            current_y += shape.height + interval_height

    def __call__(self, shapes):
        axis = self._yield_axis(self.axis, shapes)
        c_box = self._yield_circumscribed_box(shapes)
        c_shape = None  # `c_shape` encloses all the Shapes.
        for shape in shapes:
            if shape.box == c_box:
                c_shape = shape
                break
        else:
            c_shape = None

        if c_shape:
            shapes = [shape for shape in shapes if shape.Id != c_shape.Id]
            is_edge_interval = True
        else:
            is_edge_interval = False

        if axis == 0:
            self._adjust_vertially(shapes, c_box, is_edge_interval)
        elif axis == 1:
            self._adjust_horizontally(shapes, c_box, is_edge_interval)
        else:
            raise RuntimeError("Bug.")


