

from typing import Literal, assert_never, Callable, assert_never
from fairypptx.shape_range.types import AlignCMD, AlignParam
from fairypptx.shape_range import ShapeRange 
from fairypptx.box import Box 


def from_align_cmd(align_cmd: AlignCMD) -> int:
    match align_cmd:
        case "left":
            # msoAlignLeft: 0
            return 0
        case "center":
            # msoAlignCenter: 1
            return 1
        case "right":
            # msoAlignRight: 2
            return 2
        case "top":
            # msoAlignTop: 3
            return 3
        case "middle":
            # msoAlignMiddle: 4
            return 4
        case "bottom": 
            # msoAlignBottom: 5
            return 5
        case _ as unreachable:
            assert_never(unreachable)

class ShapeRangeAligner:
    def __init__(self, align_config: AlignCMD | AlignParam = AlignParam()):
        self.align_config  = align_config

    def __call__(self, shape_range: ShapeRange) -> None:
        if not isinstance(self.align_config, AlignParam):
            align_cmd = self.align_config
        else:
            align_cmd = self._to_align_cmd(shape_range, self.align_config)
        if 1 < len(shape_range):
            shape_range.api.Align(from_align_cmd(align_cmd), False)

    def _to_align_cmd(self, shape_range:ShapeRange, param: AlignParam) -> AlignCMD:
        def _param_to_cost(param: AlignParam):
            direction = param.direction
            pivot = param.pivot
            assert direction is not None
            assert pivot is not None
            boxes = [shape.box for shape in shape_range]

            def get_weight(box: Box, weight_axis: int | None = None) -> float: 
                if weight_axis == 0:
                    return box.y_length
                elif weight_axis == 1:
                    return box.x_length
                else:
                    return (box.y_length + box.x_length) / 2

            def _to_edge_cost(selector: Callable[[Box], float], stat_func: Callable, weight_axis: int | None = None) -> float:
                weights = [get_weight(box, weight_axis) for box in boxes]
                base = stat_func((selector(box) for box in boxes))
                return sum(abs(selector(box) - base) * weight for box, weight in zip(boxes, weights))

            def _to_middle_cost(axis: Literal[0, 1], weight_axis: int | None = None) -> float:
                weights = [get_weight(box, weight_axis) for box in boxes]
                base = Box.cover(boxes).center[axis]
                return sum(abs(box.center[axis] - base) * weight for box, weight in zip(boxes, weights))


            if direction == "horizontal":
                if pivot == "start":
                    return _to_edge_cost(lambda box: box.left, min, weight_axis=1)
                elif pivot == "end":
                    return _to_edge_cost(lambda box: box.right, max, weight_axis=1)
                elif pivot == "midpoint":
                    return _to_middle_cost(1, weight_axis=1)
            elif direction == "vertical":
                if pivot == "start":
                    return _to_edge_cost(lambda box: box.top, min, weight_axis=0)
                elif pivot == "end":
                    return _to_edge_cost(lambda box: box.bottom, max, weight_axis=0)
                elif pivot == "midpoint":
                    return _to_middle_cost(0, weight_axis=0)
            else:
                assert False
            assert False

        params = param.to_candidates()
        target_param = min(params, key=_param_to_cost)
        return target_param.to_align_cmd()



