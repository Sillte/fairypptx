import pytest
from PIL import Image
from fairypptx import Shape, GroupShape
from fairypptx import ShapeRange


def test_group_ungroup():
    shape1 = Shape.make(1)
    shape2 = Shape.make(2)
    sr = ShapeRange([shape1, shape2])
    p = sr.group()
    assert isinstance(p, GroupShape)
    assert len(p.children) == 2
    assert (set([shape1.api.Id, shape2.api.Id]) == 
            set([elem.api.Id for elem in p.children]))
    elems = p.ungroup()
    assert len(elems) == 2
    assert (set([shape1.api.Id, shape2.api.Id]) == 
            set([elem.api.Id for elem in elems]))



if __name__ == "__main__":
    pytest.main([__file__, "--capture=no"])
