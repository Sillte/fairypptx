from collections import defaultdict
from typing import Sequence, Self
from fairypptx.shape_range import ShapeRange 
from fairypptx.shape import Shape 
from fairypptx.box import Box 
from fairypptx.shape_range.types import AlignCMD, AlignParam
from fairypptx.shape_range.editors.aligner import ShapeRangeAligner


class UnionFind:
    def __init__(self, n: int) -> None:
        self._parent: list[int] = list(range(n)) 
        # rank: The worst-case distance to the leaf from the root.
        # The value is useful only for root nodes.  
        self._rank: list[int] = [0] * n

    def _find_root(self, x: int) -> int:
        if self._parent[x] == x:
            return x
        else:
            self._parent[x] = self._find_root(self._parent[x])
            return self._parent[x]

    def union(self, x:int, y:int) -> None:
        x, y = self._find_root(x), self._find_root(y)
        if self._rank[x] < self._rank[y]:
            x, y = y, x
        self._parent[y] = x
        if self._rank[x] == self._rank[y]:
            self._rank[x] += 1
    @property
    def groups(self) -> Sequence[Sequence[int]]:
        groups = defaultdict(list)
        for ind in range(len(self._parent)):
            groups[self._find_root(ind)].append(ind)
        return sorted(groups.values())


class ShapeCluster:
    def __init__(self, shape_range: Sequence[Shape] | ShapeRange | None = None) -> None:
        if shape_range is None:
            shape_range = ShapeRange()
        if not isinstance(shape_range, ShapeRange):
            shape_range = ShapeRange(shape_range)
        self._shape_range = shape_range
        if len(self.shape_range) == 0:
            raise ValueError("ShapeCluster requires at least one shape.")
    @property
    def shape_range(self) -> ShapeRange:
        return self._shape_range

    def expand(self, number: int = 1, candidates: Sequence[Shape] | ShapeRange | None = None) -> ShapeRange: 
        candidates = candidates or self._default_candidates() 
        candidates = [cand for cand in candidates if cand not in set(self.shape_range)]
        if len(candidates) < number:
            raise ValueError(f"Shape candidates is less than {number}.")
        for _ in range(number):
            added = self._expand_once(candidates)
            candidates.remove(added)
        return self.shape_range

    def _expand_score(self, box:Box, cluster_box: Box) -> tuple[float, float]:
        return (Box.intersection_over_cover(box, cluster_box), -Box.center_distance(box, cluster_box))

    def _expand_once(self, candidates: Sequence[Shape]) -> Shape:
        cluster_box = self.shape_range.box
        target = max(candidates, key=lambda shape: self._expand_score(shape.box, cluster_box))
        self.shape_range.append(target)
        return target

    def shrink(self, number: int = 1) -> ShapeRange:
        if len(self.shape_range) < number:
            raise ValueError(f"ClusterSize is less than {number}.")

        for _ in range(number):
            self._shrink_once()
        return self.shape_range

    def _shrink_score(self, box:Box, cluster_box: Box) -> float:
        return (Box.center_distance(box, cluster_box))

    def _shrink_once(self) -> None:
        cluster_box = self.shape_range.box
        target = max(list(self.shape_range), key=lambda shape: self._shrink_score(shape.box, cluster_box))
        self.shape_range.remove(target)

    def _default_candidates(self) -> ShapeRange:
        from fairypptx.shapes import Shapes 
        shapes = Shapes(self.shape_range)
        return shapes[:]

    @property
    def box(self) -> Box:
        return self.shape_range.box

    def append(self, shape:Shape) -> None:
        self.shape_range.append(shape)

    def align(self, align_config: AlignCMD | AlignParam = AlignParam()) -> None:
        aligner = ShapeRangeAligner(align_config = align_config)
        aligner(self.shape_range)


    @classmethod
    def from_clusters(cls, clusters:Sequence[Self]) -> Self:
        shape_range = ShapeRange.from_ranges([cluster.shape_range for cluster in clusters])
        return cls(shape_range)


class ClusterMaker:
    def __init__(self, iou_thresh: float = 0.1, background_thresh: float=0.9) -> None: 
        self.iou_thresh = iou_thresh
        self.background_thresh = background_thresh


    def __call__(self, shape_range: ShapeRange) -> Sequence[ShapeCluster]:
        backgrounds = self._extract_backgroud_shapes(shape_range)
        shape_range = ShapeRange([shape for shape in shape_range if shape not in backgrounds]) 
        clusters = [ShapeCluster([shape]) for shape in shape_range]
        clusters = self._clusters_by_ious(clusters)
        clusters = self._merge_clusters(clusters)
        if backgrounds:
            b_cluster = ShapeCluster(backgrounds) 
            return [*clusters, b_cluster]
        else:
            return clusters
    
    def _extract_backgroud_shapes(self, shape_range: ShapeRange) -> Sequence[Shape]:
        box = shape_range.box
        result = []
        for shape in shape_range:
            if Box.intersection_over_union(box, shape.box) >= self.background_thresh:
                result.append(shape)
        return result

    def _clusters_by_ious(self, clusters: Sequence[ShapeCluster]) -> Sequence[ShapeCluster]:
        N = len(clusters)
        union_find = UnionFind(N)
        for i in range(N):
            for j in range(i + 1, N):
                if Box.intersection_over_union(clusters[i].box, clusters[j].box) >= self.iou_thresh:
                    union_find.union(i, j)
        return [ShapeCluster.from_clusters([clusters[ind] for ind in group]) for group in union_find.groups]

    def _merge_clusters(self, clusters: Sequence[ShapeCluster]) -> Sequence[ShapeCluster]:
        N = len(clusters)
        union_find = UnionFind(N)
        for i in range(N):
            for j in range(i + 1, N):
                if Box.intersection_over_union(clusters[i].box, clusters[j].box) >= self.iou_thresh:
                    union_find.union(i, j)
        return [ShapeCluster.from_clusters([clusters[ind] for ind in group]) for group in union_find.groups]

