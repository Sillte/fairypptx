"""There are codes related to modify `TextRange`.

"""
from typing import Sequence, Literal
import itertools
from fairypptx.text_range import TextRange
from fairypptx.text_range.formatters import guessers
from fairypptx import constants

class TrimItemization: 
    """Itemization as for `Trim`.

    Args: 
        head_spaces: the line spaces of the start of `itemization`  
        tail_spaces: the line spaces of the end of `itemization`  
    """

    ITEMIZATION_TYPES = {constants.ppBulletUnnumbered, constants.ppBulletNumbered}
    def __init__(self, head_spaces: int =1, tail_spaces: int=1):
        self.head_spaces = head_spaces
        self.tail_spaces = tail_spaces

    def __call__(self, textrange: TextRange) -> TextRange:
        if not textrange.text.strip("\r\013"):
            return textrange
        paragraphs = textrange.paragraphs
        for para in textrange.paragraphs:
            if self._is_itemization(para):
                if self._is_empty(para):
                    para.api.Delete()
                    
        # Since `Delete` is performed, so re-getting is mandatory.
        paragraphs = textrange.paragraphs
        for para in paragraphs:
            if self._is_itemization(para):
                if not self._is_empty(para):
                    para.editor.set_tail_newlines(1)

        # Since `Delete` is performed, so re-getting is mandatory.
        paragraphs = textrange.paragraphs
        groups = [(key, list(paras)) for key, paras
                                     in itertools.groupby(paragraphs, key=self._is_itemization)
                                     if  key]
        for key, paras in reversed(groups):
            paras[0].editor.set_head_newlines(self.head_spaces + 1)
            paras[-1].editor.set_tail_newlines(self.tail_spaces + 1)
        return textrange


    def _is_itemization(self, paragraph):
        return paragraph.api.ParagraphFormat.Bullet.Type in self.ITEMIZATION_TYPES

    def _is_empty(self, paragraph):
        return not paragraph.text.rstrip("\r\013")


class LineSpacer:
    """Modify the number of line-spaces.
    Specifically, the line spaces more than `n_spaces` 
    reduces to `n_space`. 
    """

    def __init__(self, n_spaces: int =1):
        self.n_spaces = n_spaces
        pass

    def __call__(self, textrange: TextRange) -> TextRange:
        if not textrange.text.strip("\r\013"):
            return textrange

        for para in textrange.paragraphs:
            if para.editor.n_tail_newlines >= self.n_spaces + 2:
                para.editor.set_tail_newlines(self.n_spaces + 1)
        if textrange.editor.n_head_newlines >= self.n_spaces + 2:
            textrange.editor.set_head_newlines(self.n_spaces + 1)
        return textrange



class HeaderSpacer:
    """Modify the spaces related to `Header`.

    [CAUTION]
        This function is basend on `guessing`. 

    Args:
        `ignore_beginning`: The first paragraph is also modified or not.
    """
    def __init__(self,
                 head_spaces=1,
                 tail_spaces=1,
                 ignore_beginning=True):
        self.head_spaces = head_spaces
        self.tail_spaces = tail_spaces
        self.ignore_beginning = ignore_beginning


    def __call__(self, textrange: TextRange) -> TextRange:
        groups = guessers.guess_header_paragraphs(textrange)
        for level, headers in groups:
            for header in headers:
                p_index = header.paragraph_index
                if p_index == 0 and self.ignore_beginning is True:
                    pass
                else:
                    header.editor.set_head_newlines(self.head_spaces + 1)
                header.editor.set_tail_newlines(self.tail_spaces + 1)
        return textrange

class Composer:
    def __init__(self, *args):
        if len(args) == 1 and isinstance(args, Sequence): 
            args = args[0]
        self.funcs = args

    def __call__(self, target):
        for func in self.funcs:
            target = func(target)
        return target
         



class DefaultFormatter:
    """Based on authors' experience and fairies' capricious, 
    modify TextRange. 
    """

    def __call__(self, textrange):
        textrange = TextRange(textrange)
        editors = []
        editors.append(TrimItemization(head_spaces=1, tail_spaces=1))
        editors.append(HeaderSpacer(head_spaces=1, tail_spaces=1))
        editors.append(LineSpacer(n_spaces=1))
        caller = Composer(editors)

        textrange =  caller(textrange)
        textrange.shape.tighten()
        return textrange

class FontResizer:
    """Resize the font of `TextRange`.
    Args:
        pass
    """
    def __init__(self,
                 fontsize: int =18,
                 mode: Literal["min", "all"]="min"):
        self.fontsize = fontsize
        self.mode = mode

    def __call__(self, textrange: TextRange) -> TextRange:
        textrange = TextRange(textrange)
        if not textrange.text:
            raise ValueError("Empty Textrange.")

        if self.mode == "all":
            textrange = TextRange(textrange)
            textrange.font.size = self.fontsize
            return textrange
        elif self.mode == "min":
            min_size = min((run.font.size for run in textrange.runs), default=self.fontsize)
            diff_size = self.fontsize - min_size 
            for run in textrange.runs:
                run.font.size += diff_size
            return textrange
        else:
            raise RuntimeError("Implementation Error `mode`.")


if __name__ == "__main__":
    TEXT = """
# figpptx

### Introduction

**figpptx** performs conversion of artists of [matplotlib](https://matplotlib.org/) to [Shape Object (Powerpoint)](https://docs.microsoft.com/en-us/office/vba/api/powerpoint.shape). 

Suppose the situation you write a python code in order to make a presentation with PowerPoint.   
I bet many use [matplotlib](https://matplotlib.org/) (or the derivatives) for visualization.         
We have to transfer objects of matplotlib such as Figure to Slide of Powerpoint.    
It is desirable to perform this process swiftly, since we'd like to improve details of visualization based on the feel of Slide.  

I considered how to perform this chore efficiently.     
**figpptx** is written to integrate my experiments as a (somewhat) makeshift library.      


### Requirements

* Python 3.6+ (My environment is  3.8.2.)  
* Microsoft PowerPoint (My environment is Microsoft PowerPoint 2016)  
* See ``requirements.txt``.

### Install

Please clone or download this repository, and please execute below.  

```bat
python setup.py install 
```

### CAUTION!
This library uses [COM Object](https://docs.microsoft.com/en-us/windows/win32/com/the-component-object-model) for automatic operation of Powerpoint.    
Therefore, automatic operations are performed at your computer. Don't be panick!   

"""
    from fairypptx import Markdown, TextRange, Shape
    #shape = Markdown.make(TEXT).shape
    shape = Shape()
    print(shape.textrange.text)
    TrimItemization()(shape.textrange);
    print(shape.textrange.text)
    #exit(0)


    PageSpacer()(shape.textrange); 
    #groups = guessers.guess_header_paragraphs(shape.textrange)
    # print(groups)
    print(shape.text)

    DefaultFormatter()(shape.textrange)
    print("end")
