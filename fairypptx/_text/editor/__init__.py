"""There are codes related to modify `TextRange`.

"""
from fairypptx.text import TextRange

class TextRangeModifier:
    """
    """
    def __init__(self):
        pass

    def __call__(self, textrange): 
        return textrange


class MarginMaintainer:
    """Maintain the margin of `textrange`.
    Especially, it is intended to handle between `sections`.
    In addition, 
    """
    def __init__(self):
        # If you would add configurations.  
        pass

    def __call__(self, textrange):
        textrange = section_margin(textrange)
        textrange = strip_lines(textrange)
        return textrange


def strip_lines(textrange):
    """Strip the empty lines at the start and end.
    """
    targets = set()
    paragraphs = textrange.paragraphs
    for index in reversed(range(len(paragraphs))):
        content = paragraphs[index].text.strip("\n\r\v")
        if not content:
            targets.add(index)
        else:
            break
    for index in range(len(paragraphs)):
        content = paragraphs[index].text.strip("\n\r\v")
        if not content:
            targets.add(index)
        else:
            break
    for t in sorted(targets, reverse=True):
        paragraphs[t].delete()

    paragraphs = textrange.paragraphs
    last_paragraph = paragraphs[-1]
    last_runs = last_paragraph.runs
    targets = set()
    for t in reversed(range(len(last_runs))):
        if not last_runs[t].text.strip("\n\r\v"):
            last_runs[t].delete()

def get_linebreaks(textrange, mode="end"):
    """Returns the number of the `\n` and `\v` 

    Args:
        mode:
            **end**
            It returns`\n` appears how many times at the end of this `textrange`.
    """
    mode = mode.lower()
    assert mode == "end"

    count = 0
    text = textrange.text
    i = len(text) - 1
    while 0 <= i and text[i] in {"\v", "\r", "\n"}:
        i = i - 1
        count += 1
    return count


def is_empty(textrange):
    """Return whether `textrange` contains
    the visible elements or not. 
    """
    return not bool(textrange.text.strip("\v\r\n "))


def get_titles(textrange):
    """Return `Paragraphs` which are regarded as `Title`.
    For the definition of `title`, refer to `is_title`. 
    """
    def _to_min_fontsize(textrange):
        try:
            return min((run.font.size for run in textrange.runs if not is_empty(run)))
        except TypeError:
            return textrange.font.size

    def _is_all_bold(textrange):
        return all(run.font.bold for run in textrange.runs)

    def _is_larger(target, base, with_equal=True): 
        target_fontsize = _to_min_fontsize(target)
        target_bold = _is_all_bold(target)
        base_fontsize = _to_min_fontsize(base)
        base_bold = _is_all_bold(base)
        if not with_equal:
            return (target_fontsize, target_bold) > (base_fontsize, base_bold)
        else:
            return (target_fontsize, target_bold) >= (base_fontsize, base_bold)

    def _key(textrange):
        fontsize = _to_min_fontsize(textrange)
        bold = _is_all_bold(textrange)
        return (fontsize, bold)
    tr = TextRange(textrange.api.Parent.TextRange)
    pairs = []
    for para in tr.paragraphs:
        empty = is_empty(para)
        if (not empty):
            pairs.append((para, _key(para)))
    titles = []
    if len(pairs) == 0:
        return []
    elif len(pairs) == 1:
        return pairs[0]

    if pairs[0][-1] > pairs[1][-1]:
        titles.append(pairs[0][0])
    for i in range(1, len(pairs) - 1):
        if max(pairs[i-1][-1], pairs[i+1][-1]) < pairs[i][-1]:
            titles.append(pairs[i][0])
    if pairs[-1][-1] > pairs[-2][-1]:
        titles.append(pairs[-1][0])
    return titles


def is_title(textrange):
    """Check whether given `textrange` is regarded as title or not.

    Note
    ----
    * Title must contain only one `run`.  
    * Title must be `bold` or has larger fontsize than the perihperal paragraphs. 
    """

    def _to_min_fontsize(textrange):
        try:
            return min((run.font.size for run in textrange.runs if not is_empty(run)))
        except TypeError:
            return textrange.font.size

    def _is_all_bold(textrange):
        return all(run.font.bold for run in textrange.runs)

    def _is_larger(target, base, with_equal=True): 
        target_fontsize = _to_min_fontsize(target)
        target_bold = _is_all_bold(target)
        base_fontsize = _to_min_fontsize(base)
        base_bold = _is_all_bold(base)
        if not with_equal:
            return (target_fontsize, target_bold) > (base_fontsize, base_bold)
        else:
            return (target_fontsize, target_bold) >= (base_fontsize, base_bold)


    # For title, `runs` must be 1.
    runs = [run for run in textrange.runs if not is_empty(run)]
    if not runs:
        return False
    

    next_p = to_next_paragraph(textrange)
    if next_p:
        if _is_larger(next_p, textrange, with_equal=True):
            return False
    prev_p = to_prev_paragraph(textrange)
    if prev_p:
        if _is_larger(prev_p, textrange, with_equal=True):
            return False
    return True


def to_next_paragraph(textrange, with_empty=False):
    """Returns the next paragraph.
    If next paragraph does not existent,
    then `None` is returned.
    if `only_valid` is `True`, then empty paragraphs
    are not taken into consideration.
    """

    parent = TextRange(textrange.api.Parent.TextRange)
    pairs = []
    for para in parent.paragraphs:
        empty = is_empty(para)
        if with_empty is True or (not empty):
            pairs.append((para.api.Start, para))

    target = textrange.api.Start
    for c, n in zip(pairs[:-1], pairs[1:]):
        ci, _ = c
        ni, para = n
        if ci <= target < ni:
            return para
    return None

def to_prev_paragraph(textrange, with_empty=False):
    """Returns the previous paragraph.
    If previous paragraph does not existent,
    then `None` is returned.
    """
    parent = TextRange(textrange.api.Parent.TextRange)
    pairs = []
    for para in parent.paragraphs:
        empty = is_empty(para)
        if with_empty is True or (not empty):
            pairs.append((para.api.Start, para))
    target = textrange.api.Start
    for p, c in zip(pairs[:-1], pairs[1:]):
        pi, para = p
        ci, _ = c
        if pi < target <= ci:
            return para
    return None

def section_margin(textrange):
    """Give the at least oneline margins between `title`s.
    """

    titles = get_titles(textrange)
    for paragraph in titles:
        next_p = to_next_paragraph(paragraph, with_empty=True)
        if not next_p:
            continue
        if is_empty(next_p):
            continue
        breaks = get_linebreaks(paragraph)
        if breaks <= 1:
            paragraph.text += "\n"
    return textrange
