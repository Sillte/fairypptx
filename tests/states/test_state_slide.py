from fairypptx.slides import Slides
from fairypptx.states.slide import SlideStateModel
from fairypptx.states.context import Context
from fairypptx.states.shape import Shape


def test_slide_state_roundtrip():
    slide = Slides().add()

    s1 = slide.shapes.add(1)
    s1.text = "Hello"

    s2 = slide.shapes.add(1)
    s2.text = "World"

    state = SlideStateModel.from_entity(slide)
    state.apply(slide)

    assert [s.text for s in slide.shapes] == ["Hello", "World"]

def test_slide_state_create():
    slide = Slides().add()
    slide.shapes.add(1).text = "A"
    slide.shapes.add(1).text = "B"

    state = SlideStateModel.from_entity(slide)

    new_slide = state.create_entity(Context(presentation=slide.presentation))

    assert [s.text for s in new_slide.shapes] == ["A", "B"]

def test_slide_apply_id_matching_only():
    slide = Slides().add()
    s1 = slide.shapes.add(1)
    s1.text = "A"

    state = SlideStateModel.from_entity(slide)

    # shape を追加して不整合を作る
    slide.shapes.add(1).text = "B"

    # apply は落ちない（警告のみ）
    state.apply(slide)

    assert slide.shapes[0].text == "A" or  slide.shapes[1].text == "A"


def test_zorder_positions():
    def _sorted_shape_text(shape_list: list[Shape]) -> list[str]:
        return [s.text for s in sorted(shape_list, key=lambda shape: shape.api.ZOrderPosition)]
    slide = Slides().add()
    s1 = slide.shapes.add(1)
    s1.text = "A" 
    s2 = slide.shapes.add(1)
    s2.text = "B"
    texts = _sorted_shape_text([s1, s2])
    state = SlideStateModel.from_entity(slide)
    slide = Slides().add()
    context = Context(slide=slide)
    slide = state.create_entity(context)
    assert texts == _sorted_shape_text(list(slide.shapes))

def test_note_text_frame():
    slide = Slides().add()
    slide.note_text_frame.text_range.text = "HelloNote"
    state = SlideStateModel.from_entity(slide)
    slide = Slides().add()
    state.apply(slide)
    assert slide.note_text_frame.text_range.text == "HelloNote"


