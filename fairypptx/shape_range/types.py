from typing import Literal, Self, Sequence, assert_never
from dataclasses import dataclass

type AlignCMD = Literal["left", "center", "right", "top", "middle", "bottom"]
type AlignDirection = Literal["horizontal", "vertical"]
type AlignPivot = Literal["start", "midpoint", "end"]


@dataclass
class AlignParam:
    direction: AlignDirection | None = None
    pivot: AlignPivot | None = None

    def to_candidates(self) -> Sequence["AlignParam"]:
        ALL_DIRECTIONS: Sequence[AlignDirection] = ["horizontal", "vertical"]
        ALL_PIVOTS: Sequence[AlignPivot] = ["start", "midpoint", "end"]

        directions = [self.direction] if self.direction else ALL_DIRECTIONS
        pivots = [self.pivot] if self.pivot else ALL_PIVOTS
        return [AlignParam(direction=direction, pivot=pivot) for direction in directions for pivot in pivots]

    def to_align_cmd(self) -> AlignCMD:
        if self.direction is None or self.pivot is None:
            raise ValueError("AlignParam must have both 'direction' and 'pivot' defined to convert to AlignCMD.")

        match self.direction:
            case "horizontal":
                match self.pivot:
                    case "start":
                        return "left"
                    case "midpoint":
                        return "center"
                    case "end":
                        return "right"
                    case _ as unreachable:
                        assert_never(unreachable)

            case "vertical":
                match self.pivot:
                    case "start":
                        return "top"
                    case "midpoint":
                        return "middle"
                    case "end":
                        return "bottom"
                    case _ as unreachable:
                        assert_never(unreachable)

            case _ as unreachable:
                assert_never(unreachable)
