import numpy as np
from typing import List
from fairypptx.box import Box
from fairypptx.slide import Slide


class RangeIndexer:
    """This class converts `float` value (`pivot`) to `int` value (`index`.).

    The correspondence between `index` and `pivots`
    --------------
    index = -1:  (-\inf, `pivots[0]`).
    index = len(`pivots`): [`pivots[len(pivots) - 1], +\inf)
    Otherwise : The range [`pivots[`index`]` , `pivots[`index`] - 1).
    Notice the difference between closed interval and open interval.
    """

    def __init__(self, pivots: List[float], eps: float = None):
        self._pivots = self._gen_pivots(pivots, eps)
        self.eps = eps

    @property
    def pivots(self):
        """Return the pivots of the partition."""
        return self._pivots

    def __getitem__(self, target) -> int:
        assert False
        return self.to_index(target)

    def to_index(self, target) -> int:
        """Return the index of the value.
        Notice that `-1` and `len(self.pivots)`
        means the outside of the pivots.
        """
        pivots = self.pivots

        def _inner(left, right):
            if left == right:
                return left
            if right - left == 1:
                if pivots[left] <= target < pivots[right]:
                    return left
                elif pivots[right] <= target < pivots[right + 1]:
                    return right
            center = (left + right) // 2
            if pivots[center] <= target < pivots[right]:
                return _inner(center, right)
            elif pivots[left] <= target < pivots[right]:
                return _inner(left, center)

        if target < pivots[0]:
            return -1
        elif pivots[len(pivots) - 1] <= target:
            return len(pivots)
        return _inner(0, len(self.pivots) - 1)

    def _gen_pivots(self, pivots, eps):
        pivots = sorted(set(pivots))
        if eps is None:
            return pivots
        assert isinstance(eps, (float, int))
        result = list()
        base_value = pivots[0]
        result.append(base_value)
        for value in pivots[1:]:
            if value - base_value >= eps:
                result.append(base_value)
                base_value = value
        result.append(base_value)
        result = sorted(set(result))  # It is required for the corner case.
        return result


#z_values = [0, 1, 3.5]
#indexer = RangeIndexer(z_values)
#assert indexer.to_index(2.0) == 1
#assert indexer.to_index(-45) == -1


class GridHandler:
    """Handling of `grid` of Slide.

    Here, `Grid` is a collection of rectangle tiles.
    these tiles are possible to access with indices [x_index, y_index].
    via `RangeIndexer`.

    Attributes:
        x_indexer: `RangeIndexer`
        y_indexer: `RangeIndexer`.
            They manage the relationship between index and the pivots(pixel values).
            Refer to `RangeIndexer`, for details.
        occupations: np.ndarray(bool). (len(y_indexer.pivots) - 1, len(x_indexer.pivots) - 1)
            It represents whether `Shape` exist in grid (yi, xi).
            The index (yi, xi) corresponds to the range
            ( `x_indexer.pivots[xi],x_indexer.pivots[xi + 1]) \times
            ( `y_indexer.pivots[yi],y_indexer.pivots[yi + 1]).

    Note
    -------------------------------------
    * Objects outside `Slide` is ignored.
    """

    def __init__(self, slide=None):
        self.slide = Slide(slide)
        self.x_indexer, self.y_indexer = self._make_grids(self.slide)
        self.occupations = self.make_occupations(self.slide.shapes)

    def _make_grids(self, slide):
        x_pivots = []
        y_pivots = []
        shapes = slide.shapes
        for shape in slide.shapes:
            y_pivots.append(shape.top)
            y_pivots.append(shape.top + shape.height)
            x_pivots.append(shape.left)
            x_pivots.append(shape.width + shape.left)
        x_pivots.extend([0.0, slide.width])
        y_pivots.extend([0.0, slide.height])

        # Ignore outsize of the slider.
        x_pivots = [elem for elem in x_pivots if 0 <= elem <= slide.width]
        y_pivots = [elem for elem in y_pivots if 0 <= elem <= slide.height]

        x_indexer = RangeIndexer(x_pivots)
        y_indexer = RangeIndexer(y_pivots)
        return x_indexer, y_indexer

    def make_occupations(self, shapes):
        """Return the `occupations`.

        Args:
            shapes (Shapes):
            If any of these exist, the grid is considered to be occupied.

        Return:
            occupations (np.ndarray) two-dimension,
            represents whether its grid is occuped or not.

        """
        assert self.x_indexer
        assert self.y_indexer
        xn = len(self.x_indexer.pivots)
        yn = len(self.y_indexer.pivots)
        occupations = np.zeros((yn - 1, xn - 1), np.bool)

        for shape in shapes:
            sx, ex = shape.left, shape.left + shape.width
            sy, ey = shape.top, shape.top + shape.height
            x_si = self.x_indexer.to_index(sx)
            x_ei = self.x_indexer.to_index(ex)
            y_si = self.y_indexer.to_index(sy)
            y_ei = self.y_indexer.to_index(ey)
            occupations[y_si:y_ei, x_si:x_ei] = True
        return occupations

    def make_tiles(
        self, true_color=(255, 0, 255, 128), false_color=(128, 128, 128, 128)
    ):
        """Making tiles inside `grids`.
        Note
        ---------------
        This is mainly for debug.
        """
        x_pivots = self.x_indexer.pivots
        y_pivots = self.y_indexer.pivots
        for sx, ex in zip(x_pivots[:-1], x_pivots[1:]):
            for sy, ey in zip(y_pivots[:-1], y_pivots[1:]):
                shape = Shape.make(1)
                x_si = self.x_indexer.to_index(sx)
                y_si = self.y_indexer.to_index(sy)
                if self.occupations[y_si, x_si]:
                    color = true_color
                else:
                    color = false_color
                shape.top = sy
                shape.height = ey - sy
                shape.left = sx
                shape.width = ex - sx
                shape.fill = color

    def get_maximum_box(self, occupations=None) -> Box:
        """Return the maximum empty `Box`.

        Args:
            occupations (np.ndarray) two dimension.:
            It represents whether the shape exists in the specified grid or not.
        """
        if occupations is None:
            occupations = self.occupations

        y_pivots = self.y_indexer.pivots
        x_pivots = self.x_indexer.pivots

        # Firstly, generate lengh table
        # dist: (si, ei) -> float.
        # the distance from `si` to `ei`,  inclusive.
        # si <= ei.
        def _gen_length_table(pivots):
            result = dict()
            n_pivot = len(pivots)
            for si in range(n_pivot - 1):
                d = 0
                for ei in range(si, n_pivot - 1):
                    d += pivots[ei + 1] - pivots[ei]
                    result[(si, ei)] = d
            assert len(result) == (len(pivots)) * (len(pivots) - 1) / 2
            return result

        y_dist = _gen_length_table(y_pivots)
        x_dist = _gen_length_table(x_pivots)

        # Here, the maximum bleak area is calculated
        # via naive implementation.
        # (2021/04/29) It is better to rewrite it
        # to more sophisticated ones.
        def _to_maximum(yi, xi):
            if occupations[yi, xi] == True:
                return None, None
            m_area = 0
            result = None
            for yu in range(yi + 1):
                for xu in range(xi + 1):
                    if np.all(np.logical_not(occupations[yu : yi + 1, xu : xi + 1])):
                        area = y_dist[(yu, yi)] * x_dist[(xu, xi)]
                        if m_area < area:
                            m_area = area
                            result = (yu, xu)
            return result, m_area

        yn, xn = occupations.shape
        data = dict()
        for yi in range(yn):
            for xi in range(xn):
                pair, area = _to_maximum(yi, xi)
                if pair is None:
                    continue
                key = (*pair, *(yi, xi))
                data[(*pair, *(yi, xi))] = area
        ys, xs, ye, xe = max(data, key=lambda k: data[k])
        xs = x_pivots[xs]
        xe = x_pivots[xe + 1]

        ys = y_pivots[ys]
        ye = y_pivots[ye + 1]

        left = xs
        width = xe - xs
        top = ys
        height = ye - ys
        box = Box(left=left, top=top, width=width, height=height)
        return box


if __name__ == "__main__":
    pass
