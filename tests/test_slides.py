# tests/test_slides_integration.py
import pytest
from fairypptx.presentation import Presentation
from fairypptx.slides import Slides   # ← いま書いてくれたクラス
from fairypptx.slide import Slide
from fairypptx.slide_range import SlideRange


@pytest.fixture
def pres():
    pres = Presentation()
    while pres.slides.api.Count > 0:
        pres.slides.api.Item(1).Delete()
    yield pres



def test_slides_len_and_add(pres):
    slides = Slides(pres.api.Slides)

    assert len(slides) == 0

    s1 = slides.add()
    assert isinstance(s1, Slide)
    assert len(slides) == 1

    s2 = slides.add()
    assert s2.api.SlideIndex == 2
    assert len(slides) == 2

# ---------------------------------------------
# __getitem__ with int
# ---------------------------------------------
def test_slides_getitem_int(pres):
    slides = Slides(pres.api.Slides)
    s1 = slides.add()
    s2 = slides.add()

    s = slides[1]
    assert isinstance(s, Slide)
    assert s.api.SlideIndex == 2


# ---------------------------------------------
# __getitem__ with slice → SlideRange
# ---------------------------------------------
def test_slides_getitem_slice(pres):
    slides = Slides(pres.api.Slides)

    s1 = slides.add()
    s2 = slides.add()
    s3 = slides.add()

    sr = slides[0:2]  # first two slides

    assert isinstance(sr, SlideRange)
    assert len(sr) == 2
    assert sr[0].api.SlideIndex == 1
    assert sr[1].api.SlideIndex == 2


def test_slides_iter(pres):
    slides = Slides(pres.api.Slides)
    s1 = slides.add()
    s2 = slides.add()

    collected = list(slides)
    assert len(collected) == 2
    assert all(isinstance(s, Slide) for s in collected)
    assert collected[0].api.SlideIndex == 1
    assert collected[1].api.SlideIndex == 2
