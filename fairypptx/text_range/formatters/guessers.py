"""Guess the information from `TextRange`. 

Situations where we would like to extract information  
from `TextRange`. 
The functions or classes in this module for these situations. 


Rule
----
* We do not modify or change given `TextRange` here. 


Fuzzy Definitions for terminology. 
(If anything desirable to be mentioned, )

"""
from typing import Sequence, Dict, List, Tuple
import itertools
from collections import defaultdict
from dataclasses import dataclass



def guess_default_fontsize(textrange) -> float:
    """Return the fontsize which is considered 
    `default` inside `textrange`. 

    Policy:
    --------
    * The font must be often used. 
    * The font must be smaller than the one used for emphasizing. 
    """
    if not textrange.text:
        raise ValueError("Empty TextRange.")

    paragraphs = textrange.paragraphs

    size_to_npara = defaultdict(int)

    for para in paragraphs:
        sizes = [run.font.size for run in para.runs]
        if len(set(sizes)) == 1:
            size_to_npara[sizes[0]] += 1
    if not size_to_npara:
        default_fontsize = paragraphs[-1].font.size  # Fallback.
    else:
        # outliers' removal 
        threshold = sorted(size_to_npara.values())[min(round(len(size_to_npara) * 1 // 4), len(size_to_npara) -1)]
        size_to_npara = {key: value for key, value in size_to_npara.items() if threshold <= value}
        default_fontsize = min(size_to_npara.keys())
    return default_fontsize

def guess_header_paragraphs(textrange) -> List[Tuple[int, List["TextRange"]]]:
    """Return the `list` whose element is
    (`indent_level`, `paragraphs`). 
    Here, `indent_level` starts from 0.

    Murmurs:
        The information used for `header`'s guessing. 

        * Their Bullet type is not `ITEMIZATION_TYPES`  
        * `fontsize` must be larger than or equal to `normal`  
        * If `fontsize` is the same as `normal`, `bold` or `underline` is applied over all the paragraphs.
    """
    paragraphs = textrange.paragraphs

    default_fontsize = guess_default_fontsize(textrange)
    prop = _gen_fontsize_properties(textrange)[default_fontsize]

    is_bold_dominant = (prop.bolds / prop.total > 0.5)
    is_underline_dominant = (prop.underlines / prop.total > 0.5)

    def _is_header(paragraph):
        if paragraph.api.ParagraphFormat.Bullet.Visible:
            return False
        if _is_empty(paragraph):
            return False

        fontsize = _to_single_fontsize(paragraph)

        if fontsize < default_fontsize:
            return False
        elif fontsize == default_fontsize:
            is_bold_cond = ((not is_bold_dominant) and all(run.font.bold for run in _runs_iter(paragraph)))
            is_underline_cond = ((not is_underline_dominant) and all(run.font.underline for run in _runs_iter(paragraph)))
            return (is_bold_cond or is_underline_cond)
        else:
            return True

    # Firstly, judge whether `each` paragraph is `Header` or not.
    # Secondly, determine the `level` of `header`. 

    def key_func(paragraph):
        fontsize = _to_single_fontsize(paragraph)
        is_underline = paragraph.font.underline
        is_bold = paragraph.font.bold
        return (fontsize, is_underline, is_bold)

    headers = [para for para in paragraphs if _is_header(para)]
    headers = sorted(headers, key=key_func, reverse=True)
    pairs   = [(level, list(values)) for level, (key, values)
                         in enumerate(itertools.groupby(headers, key=key_func))]
    return pairs




@dataclass
class _PropertyCounter:
    """Represents the number of `Characters`
    for each property of `TextRange`. 
    The unit is `characters`.  
    """
    bolds: int = 0 
    underlines:int = 0
    total: int = 0


def _gen_fontsize_properties(textrange):
    """Return the dict. Key is `fontsize` and  
    the value is also `dict`, which represents 
    the information about `TextRange`s for each `fontsize`.  
    """
    result = defaultdict(_PropertyCounter)
    for run in textrange.runs:
        fontsize = run.font.size
        counter = result[fontsize]
        counter.total += run.api.Length
        if run.font.bold:
            counter.bolds += run.api.Length
        if run.font.underline:
            counter.underlines += run.api.Length
    return result

def _is_empty(textrange):
    """Here `empty` means some characters exists except
    `IGNORE_CHARS` (e.g. `\r`)
    """
    IGNORE_CHARS = {"\r", "\013"}
    return not bool(textrange.text.strip("".join(IGNORE_CHARS)))


def _runs_iter(textrange, ignore_empty=True):
    """Iterate `runs`. If `ignore_empty` is True`, then 
    `empty` textrange is omitted.
    """ 
    if not ignore_empty:
        return iter(textrange.runs) 
    return (run for run in textrange.runs if not _is_empty(run))

def _to_single_fontsize(textrange):
    """Since, `textrange` contains multiple `runs` 
    the size of `fontsize` 
    """
    if not textrange.text:
        raise ValueError("Empty TextRange")
    fontsize = textrange.runs[0].font.size
    result = min((run.font.size for run in _runs_iter(textrange)), default=fontsize)
    return result 



if __name__ == "__main__":
    from fairypptx import Markdown, TextRange
    textrange = TextRange()
    #print(_gen_fontsize_properties(Markdown().shape.textrange))
    #print(_gen_fontsize_properties(Markdown().shape.textrange))
    groups = guess_header_paragraphs(textrange)
    for level, headers in groups:
        for header in headers:
            print(header.text)
 
            
