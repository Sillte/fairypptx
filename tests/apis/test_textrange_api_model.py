from fairypptx.apis.text_range import TextRangeApiModel


def test_text_range_simple_paragraphs():
    from fairypptx import Shape
    shape = Shape.make(1)
    text_range = shape.text_range
    text_range.text = "Hello\rWorld"
    assert text_range.text == "Hello\rWorld"
    api_model = TextRangeApiModel.from_api(text_range.api)
    assert len(api_model.paragraphs) == 2

    shape = Shape.make(1)
    text_range2 = shape.text_range
    api_model.apply_api(text_range2.api)
    assert text_range2.text == "Hello\rWorld"


def test_text_range_empty_paragraphs():
    from fairypptx import Shape
    shape = Shape.make(1)
    text_range = shape.text_range
    text_range.text = "Hello\r\rWorld"
    assert shape.text_range.text == "Hello\r\rWorld"
    # `\n\r` is converted to `\r`.
    api_model = TextRangeApiModel.from_api(text_range.api)
    assert len(api_model.paragraphs) == 3
    api_model.apply_api(shape.text_range.api)
    assert shape.text_range.text == "Hello\r\rWorld"

    shape.text_range.text = "Hello\r\r\r\rWorld"
    assert shape.text_range.text == "Hello\r\r\r\rWorld"

    api_model = TextRangeApiModel.from_api(shape.text_range.api)
    assert len(api_model.paragraphs) == 5
    api_model.apply_api(shape.text_range.api)
    assert shape.text_range.text == "Hello\r\r\r\rWorld"


def test_text_range_interval_lfs():
    from fairypptx import Shape
    shape = Shape.make(1)
    text_range = shape.text_range
    text_range.text = "Hello\rWo\nrld"
    assert text_range.text == "Hello\rWo\nrld"
    api_model = TextRangeApiModel.from_api(text_range.api)
    assert len(api_model.paragraphs) == 2

    api_model.apply_api(text_range.api)
    assert text_range.text == "Hello\rWo\nrld"

    shape = Shape.make(1)
    text_range = shape.text_range
    text_range.text = "Hello\nWorld"
    assert text_range.text == "Hello\nWorld"
    api_model = TextRangeApiModel.from_api(text_range.api)
    assert len(api_model.paragraphs) == 1

    api_model.apply_api(text_range.api)
    assert text_range.text == "Hello\nWorld"
    
    
    shape = Shape.make(1)
    text_range = shape.text_range
    # `\n\r` is normalized to `\r`.
    text_range.text = "Hello\n\rWorld"
    assert text_range.text == "Hello\rWorld"
    api_model = TextRangeApiModel.from_api(text_range.api)
    assert len(api_model.paragraphs) == 2
    api_model.apply_api(text_range.api)
    assert text_range.text == "Hello\rWorld"

def test_text_range_object():
    from fairypptx import Shape
    shape = Shape.make(1)
    text_range = shape.text_range
    text_range.text = "Hello\rWorld"
    assert text_range.text == "Hello\rWorld"
    api_model = TextRangeApiModel.from_api(text_range.api)

    text_range.text = "Dummy"
    assert text_range.text == "Dummy"

    api_model.apply_api(text_range.api)
    assert text_range.text == "Hello\rWorld"


