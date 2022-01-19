"""Handling processing related to location and position. 

"""

from typing import Union
import numpy as np
from collections import defaultdict
from fairypptx.box import Box, intersection_over_cover
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
            raise ValueError(f"Invalid mode `{mode}`.")
        return arg 

    def _to_shapes(self, arg):
        """Convert to `Shapes`."""
        from fairypptx.shape import Shape, Shapes
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
        c_box = Box(c_left, c_top, c_right - c_left, c_bottom - c_top)
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


class AlignMode:
    """Specify the mode used for aligning.
    Here,
    * (0, start, left, top): The starting edge position is aligned.
    * (1, center, right, bottom): The ending edge position is aligned.
    * (0.5, end, center, middle): The center position is aligned.
    """

    START = "start"
    CENTER = "center"
    END = "end"

    def __init__(self, mode):
        self._mode = self._to_mode(mode)

    @property
    def mode(self):
        return self._mode

    def is_start(self):
        return self.mode == AlignMode.START

    def is_center(self):
        return self.mode == AlignMode.CENTER

    def is_end(self):
        return self.mode == AlignMode.END

    def _to_mode(self, arg):
        if isinstance(arg, AlignMode):
            return arg.mode
        elif isinstance(arg, (float, int)):
            if arg == 0:
                return AlignMode.START
            elif arg == 0.5:
                return AlignMode.CENTER
            elif arg == 1:
                return AlignMode.END
        elif isinstance(arg, str):
            arg = arg.lower().strip()
            if arg in {"start", "left", "top"}:
                return AlignMode.START
            elif arg in {"half", "center", "middle"}:
                return AlignMode.CENTER
            elif arg in {"end", "right", "bottom"}:
                return AlignMode.END
        raise ValueError("Cannot convert to Mode.", arg)

    @staticmethod
    def __call__(arg) -> Union["AlignMode.START", "AlignMode.CENTER", "AlignMode.END"]:
        if isinstance(arg, (float, int)):
            if arg == 0:
                return AlignMode.START
            elif arg == 0.5:
                return AlignMode.CENTER
            elif arg == 1:
                return AlignMode.END
        elif isinstance(arg, str):
            pass

    def __eq__(self, other):
        return self.mode == AlignMode(other).mode


class ShapesAligner:
    """Align Shapes.

    Args:
        `axis`; the direction of align.
        * `axis`: 0(height, y, horizontally) -> horizontally.
        * `axis`: 1(width, x, width) -> vertically..
        `mode`: The mode which specifies which edge is aligned.
            - start
            - center
            - end
    """

    def __init__(self, axis=None, mode=None):
        self.axis = axis
        self.mode = mode

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
                axis = 1
            else:
                axis = 0
        assert axis in {0, 1}
        return axis

    def _yield_mode(self, mode=None):
        if mode is None:
            mode = "start"
        return AlignMode(mode)

    def _align_vertically(self, shapes, mode):
        boxes = [shape.box for shape in shapes]
        if mode.is_start():
            top = min(box.top for box in boxes)
            for shape in shapes:
                shape.api.Top = top
        elif mode.is_end():
            bottom = max(box.bottom for box in boxes)
            for shape in shapes:
                shape.api.Top = bottom - shape.api.Height
        elif mode.is_center():
            center_y = sum(box.center[0] for box in boxes) / len(boxes)
            for shape in shapes:
                shape.api.Top = center_y - shape.api.Height / 2
        else:
            raise RuntimeError("Bug.")

    def _align_horizontally(self, shapes, mode):
        boxes = [shape.box for shape in shapes]
        if mode.is_start():
            left = min(box.left for box in boxes)
            for shape in shapes:
                shape.api.Left = left
        elif mode.is_end():
            right = max(box.right for box in boxes)
            for shape in shapes:
                shape.api.Left = right - shape.api.Width
        elif mode.is_center():
            center_x = sum(box.center[1] for box in boxes) / len(boxes)
            for shape in shapes:
                shape.api.Left = center_x - shape.api.Width / 2
        else:
            raise RuntimeError("Bug.")

    def __call__(self, shapes):
        axis = self._yield_axis(self.axis, shapes)
        mode = self._yield_mode(self.mode)

        if len(shapes) <= 1:
            return

        if axis == 0:
            self._align_vertically(shapes, mode)
        elif axis == 1:
            self._align_horizontally(shapes, mode)
        else:
            raise RuntimeError("Bug.")
        return shapes


class ClusterMaker:
    def __init__(self, axis=None, iou_thresh=0.1):
        self.axis = axis
        self.iou_thresh = iou_thresh

    def _yield_axis(self, axis, shapes):
        if axis == "width":
            axis = 1
        if axis == "height":
            axis = 0
        if axis is None:
            clusters_0 = self._cluster(shapes, axis=0)
            clusters_1 = self._cluster(shapes, axis=1)
            if len(clusters_0) <= len(clusters_1):
                axis = 0
            else:
                axis = 1
        assert axis in {0, 1}
        return axis

    @classmethod
    def suggest_axis(cls, shapes, iou_thresh=0.1):
        """Suggest the appropriate axis.
        Typically, as the number of clusters is smaller, better.
        """
        clusters_0 = cls(axis=0, iou_thresh=iou_thresh)(shapes)
        clusters_1 = cls(axis=1, iou_thresh=iou_thresh)(shapes)
        if len(clusters_0) <= len(clusters_1):
            return 0
        else:
            return 1

    def _cluster(self, shapes, axis=0):
        assert axis in {0, 1}
        c_axis = 0 if axis == 1 else 1

        def _is_same_cluster(shape1, shape2):
            box1 = shape1.box
            box2 = shape2.box
            rate = intersection_over_cover(box1, box2, axis=c_axis)
            return self.iou_thresh <= rate

        def sort_key(shape):
            if axis == 0:
                return shape.box.left
            else:
                return shape.box.top

        shapes = sorted(shapes, key=sort_key)
        boxes = [shape.box for shape in shapes]
        selected = set()
        g_indices = defaultdict(list)
        n_shape = len(shapes)
        for i1 in range(n_shape):
            if i1 in selected:
                continue
            selected.add(i1)
            g_indices[i1].append(i1)
            for i2 in range(i1 + 1, n_shape):
                if i2 in selected:
                    continue
                if _is_same_cluster(shapes[i1], shapes[i2]):
                    selected.add(i2)
                    g_indices[i1].append(i2)
        keys = sorted(g_indices, key=lambda i: sort_key(shapes[i]))
        from fairypptx import Shapes  # To thwart cyclic reference.

        clusters = [Shapes([shapes[i] for i in g_indices[key]]) for key in keys]
        return clusters

    def __call__(self, shapes):
        axis = self._yield_axis(self.axis, shapes)
        return self._cluster(shapes, axis)


class ClusterAligner:
    def __init__(self, axis=None, mode=None, iou_thresh=0.10):
        self.axis = axis
        self.mode = mode
        self.iou_thresh = iou_thresh

    def _yield_axis(self, axis, shapes):
        if axis == "width":
            axis = 1
        if axis == "height":
            axis = 0
        if axis is None:
            # `c_axis` should be the one whose number of gropu is smaller.
            c_axis = ClusterMaker.suggest_axis(shapes, iou_thresh=self.iou_thresh)
            axis = 0 if c_axis == 1 else 1
        assert axis in {0, 1}
        return axis

    def __call__(self, shapes):
        axis = self._yield_axis(self.axis, shapes)
        c_axis = 0 if axis == 1 else 1
        aligner = ShapesAligner(axis=axis, mode=self.mode)
        cluster_maker = ClusterMaker(axis=c_axis, iou_thresh=self.iou_thresh)
        clusters = cluster_maker(shapes)
        for cluster in clusters:
            aligner(cluster)
        return clusters


class ShapesArranger:
    def __init__(self, *, axis=0, margin=0.5):
        self.margin = margin
        self.axis = axis

    def __call__(self, shapes):
        if self.axis == 0:
            return self._vertial_arrange(shapes)
        elif self.axis == 1:
            return self._horizontal_arrange(shapes)
        else:
            raise NotImplementedError("Invalid.`axis`.")

    def _vertial_arrange(self, shapes):
        box = self.circumscribed_box(shapes)
        left, top = box.left, box.top
        heights = [shape.height for shape in shapes]
        y = top
        for shape in shapes:
            margin = self._to_margin(self.margin, shape, axis=0)
            shape.left = left
            shape.top = y
            y += shape.height + margin
        return shapes

    def _horizontal_arrange(self, shapes):
        box = self.circumscribed_box(shapes)
        left, top = box.left, box.top
        heights = [shape.height for shape in shapes]
        x = left
        for shape in shapes:
            margin = self._to_margin(self.margin, shape, axis=1)
            shape.left = x
            shape.top = top
            x += shape.width + margin
        return shapes

    def _to_margin(self, margin, shape, axis=0):
        """Return the pixel margin.
        If margin is less than 5, it is regarded as ratio of the length of shape.
        Otherwise, it is regarded as the pixels.
        """
        assert axis in {0, 1}
        if 5.0 <= margin:
            return margin
        else:
            if axis == 0:
                return shape.height * margin
            else:
                return shape.width * margin

    def _to_axis(self, axis):
        if axis == "width":
            axis = 1
        elif axis == "height":
            axis = 0
        assert axis in {0, 1}
        return axis

    @classmethod
    def circumscribed_box(cls, shapes):
        boxes = [shape.box for shape in shapes]
        c_left = min(box.left for box in boxes)
        c_top = min(box.top for box in boxes)
        c_right = max(box.right for box in boxes)
        c_bottom = max(box.bottom for box in boxes)
        c_box = Box(c_left, c_top, c_right - c_left, c_bottom - c_top)
        return c_box
