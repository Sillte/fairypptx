import re
import pytest
from fairypptx import TextRange
from fairypptx import Color
from fairypptx import Shape
from fairypptx import constants


@pytest.mark.parametrize(
    "mode, s1, s2, expected",
    [
        ("after", "First", "Suffix", "FirstSuffix"),
        ("before", "First", "Prefix", "PrefixFirst"),
    ],
)
def test_insert_mode(mode, s1, s2, expected):
    shape = Shape.make(constants.msoShapeRectangle)
    shape.text = s1
    tr = TextRange(shape)
    assert len(tr.lines) == 1
    assert len(tr.runs) == 1
    ret = tr.insert(s2, mode)
    assert ret.text == s2
    # Change the format
    ret.api.Font.Size = shape.TextFrame.TextRange.Font.Size + 12
    shape.tighten()
    tr = TextRange(shape)
    assert len(tr.runs) == 2
    # Paragraphs's is not changed.
    assert len(tr.paragraphs) == 1
    assert shape.api.TextFrame.TextRange.Text == expected


def test_paragraph():
    """
    Here, we would like to clarify the Paragraph's specification.

    """
    shape = Shape.make(constants.msoShapeRectangle)
    shape.text = "Line1\nLine2"
    tr = TextRange(shape)
    assert len(tr.paragraphs) == 1
    assert len(tr.runs) == 1
    assert len(tr.sentences) == 1

    """`\r\n` is regarded as separator of Paragraph and Sentence.
    """
    target = "Paragraph1\r\nParagraph2"
    shape.text = target
    tr = TextRange(shape)

    assert len(tr.paragraphs) == 2
    # Even if multiple paragraphs exist, `runs` becomes 1.
    assert len(tr.runs) == 1
    # However, sentences are separated by ``\r\n``.
    assert len(tr.sentences) == 2
    assert tr.Text == target.replace("\r\n", "\r")

    """`\r` is also regarded as separator of Paragraph and Sentence.

    POINT1.
    When the first is `\r`, then text's last is the `\r` is stripped.
    However, the first is not `\r`, the last one is not stripped.  

    POINT2.
    The first `\r` (`\r\n`) is regarded as a separator of paragraph, 
    while the last one is not regarded as separator.

    """
    target = "\rParagraph1\rParagraph2\r"
    shape.text = target
    tr = TextRange(shape)
    assert len(tr.paragraphs) == 3
    assert tr.Text == target.replace("\r\n", "\r").rstrip("\r")
    target = "Paragraph1\rParagraph2\r"
    shape.text = target
    tr = TextRange(shape)
    assert len(tr.paragraphs) == 2
    assert len(tr.sentences) == 2
    assert tr.Text == target.replace("\r\n", "\r")

    """ Example of `insert`.
    Contrary to the direct specification,
    the last `\r` is not stripped.
    """
    shape.text = ""
    tr = TextRange(shape)
    target = "\rParagraph1\rParagraph2\r"
    tr.insert(target)
    tr = TextRange(shape)
    assert len(tr.paragraphs) == 3
    assert tr.Text == target.replace("\r\n", "\r")


def test_properties():
    """
    * Characters
    * Words
    * Runs
    * Sentences
    * Paragraphs

    Strict specification of logic of division is difficult to us...
    """

    shape = Shape.make(constants.msoShapeRectangle)
    shape.text = ""
    tr = TextRange(shape)

    # Characters / Words
    sample = "two  word"
    shape.text = sample
    # Change is not transmitted.
    assert tr.Text != sample
    tr = TextRange(shape)
    assert tr.Text == sample
    assert len(tr.characters) == len(sample)
    words = re.split(r"\s+", sample)
    """ Specification of TextRange.Words() is diffcult to pursue perfectionism.
    For example, handling of period.
    """
    assert len(tr.words) == len(words)
    for elem, word in zip(tr.words, words):
        assert elem.Text.strip() == word
    """ Concatenation of elements is equivalent to the orignal one. 
    """
    assert "".join((elem.Text for elem in tr.words)) == sample

    # Sentence
    sample = "Sentence1; 'This is a pen.'.\nSentence2; 'Is this a Pen?'."
    tokens = ["Sentence1; 'This is a pen.'.\n", "Sentence2; 'Is this a Pen?'."]
    shape.text = sample
    shape.tighten()
    tr = TextRange(shape)
    assert len(tr.sentences) == len(tokens)
    for elem, token in zip(tr.sentences, tokens):
        assert elem.Text == token
    assert "".join((elem.Text for elem in tr.sentences)) == sample

    # Lines
    sample = "\n".join(["Line1", "Line2"])
    shape.text = sample
    shape.tighten()
    tr = TextRange(shape)
    assert len(tr.lines) == 2

    sample = "\r\n".join(["Line1", "Line2"])
    shape.text = sample
    shape.tighten()
    tr = TextRange(shape)
    assert len(tr.lines) == 2

    # Runs
    shape.text = "Token1."
    tr = TextRange(shape)
    t = tr.insert("\n Token2")
    t.api.Font.Size = tr.api.Font.Size + 20
    tr = TextRange(shape)
    t = tr.insert("Token3")
    t.api.Font.Size = tr.api.Font.Size - 10
    tr = TextRange(shape)
    assert len(tr.runs) == 3
    shape.tighten()


def test_font_access():
    shape = Shape.make(constants.msoShapeRectangle)
    shape.text = "Test Sentence."
    tr = TextRange(shape)
    tr.Italic = constants.msoTrue
    assert tr.api.Font.Italic == constants.msoTrue

    tr.Bold = constants.msoTrue
    assert tr.api.Font.Bold == constants.msoTrue
    shape.tighten()


def test_font():
    shape = Shape.make(constants.msoShapeRectangle)
    shape.text = "TestFont"
    tr = TextRange(shape)
    tr.api.Font.Size = 17
    font = tr.font
    assert font["Size"] == 17
    font["Size"] = 15
    tr.font = font
    assert tr.api.Font.Size == 15

    # Setting method 1
    # Directory set it using `Object` instance.
    shape.textrange.font.api.Color.RGB = 1
    assert shape.textrange.font.color == Color(1)

    # Setting method2 using `ObjectMixinDict`.
    shape.textrange.font["Color.RGB"] = 3
    assert shape.textrange.font.color == Color(3)

    shape.textrange.font.color = 4
    assert shape.textrange.font["Color.RGB"] == 4


def test_paragraphformat():
    tr = TextRange.make_itemization(["P-ITEM1", "P-ITEM2", "P-ITEM3"])
    tr.api.ParagraphFormat.Bullet.Type = constants.ppBulletUnnumbered
    ph = tr.paragraphformat
    assert ph["Bullet.Type"] == constants.ppBulletUnnumbered
    ph["Bullet.Type"] = constants.ppBulletNumbered
    tr.paragraphformat = ph
    assert tr.api.ParagraphFormat.Bullet.Type == constants.ppBulletNumbered
    tr.shape.tighten()


def test_make_itemization():
    tr = TextRange.make_itemization(["ITEM1", "ITEM2", "ITEM3"])
    assert len(tr.paragraphs) == 3
    assert tr.api.ParagraphFormat.Bullet.Type == True
    # Access is possible.
    tr.paragraphs[-1].api.IndentLevel = 2
    assert hasattr(tr, "api")
    tr.shape.tighten()


def test_find():
    tr = TextRange.make("ITEM1-ITEM2-ITEM3")
    result = tr.find("ITEM")
    assert len(result) == 3
    assert all(elem.text == "ITEM" for elem in result)


if __name__ == "__main__":
    pytest.main([__file__, "--capture=no"])
