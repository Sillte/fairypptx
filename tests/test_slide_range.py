# tests/test_slide_range_integration.py
import pytest
from fairypptx.slide import Slide
from fairypptx.slide_range import SlideRange
from fairypptx.presentation import Presentation

@pytest.fixture
def pres():
    pres = Presentation()
    while pres.slides.api.Count > 0:
        pres.slides.api.Item(1).Delete()
    yield pres

def test_slide_range_build_from_com(pres):
    s1 = pres.slides.add()
    s2 = pres.slides.add()

    api_range = pres.slides.api.Range([1, 2])
    sr = SlideRange(api_range)

    assert len(sr) == 2
    assert isinstance(sr[0], Slide)
    assert sr[0].api.SlideIndex == 1

def test_slide_range_slice(pres):
    s1 = pres.slides.add()
    s2 = pres.slides.add()
    s3 = pres.slides.add()

    sr = SlideRange([s1, s2, s3])

    sub = sr[1:]   # slice
    assert isinstance(sub, SlideRange) 
    assert len(sub) == 2
    assert sub[0].api.SlideIndex == 2
    assert sub[1].api.SlideIndex == 3

def test_slide_range_api_reconstruct(pres):
    s1 = pres.slides.add()
    s2 = pres.slides.add()

    sr = SlideRange([s1, s2])
    api = sr.api

    assert len(sr) == 2
    assert sr[0].api.SlideIndex == 1
    assert sr[1].api.SlideIndex == 2
    assert api.Count == 2
    assert api.Item(1).SlideIndex == 1
    assert api.Item(2).SlideIndex == 2

def test_slide_range_empty_api_error():
    sr = SlideRange([])
    with pytest.raises(ValueError):
        _ = sr.api
