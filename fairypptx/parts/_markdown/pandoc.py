"""Handle Pandoc.

Requirements
-----------------
[Pandoc](https://pandoc.org/#) must be installed,
and PATH environement is appropriately set
so that `pandoc --help` is executed from commandline. 


Comment 
-------------------
Though, `css` is handled in functions of this module. 
It is mainly intended to that this script is used for use-cases other than at POWERPOINT.   
This is because copy of HTML to Powerpoint does not reflect CSS sufficiently.   


Necessity of premailer
-----------------------

Maybe `premailer` is not necessary for `PowerPoint`, 
however, for e-mails this function may be necessary. 
* See https://github.com/premailer/premailer

"""

from io import StringIO
from functools import partial
from typing import Union, Optional
from pathlib import Path
import premailer
import subprocess

_this_folder = Path(__file__).absolute().parent
_default_css_cache = _this_folder / "__css__"
_default_css_cache.mkdir(exist_ok=True)


def _is_existent_path(arg):
    try:
        return Path(arg).exists()
    except:
        return False


def to_html(
    markdown,
    css: Optional[Union[str, Path]] = None,
    *,
    output_path=None,
    css_folder=None,
):
    """Convert markdown document to html."""
    css_path = _CSSCache(css_folder)(css)
    command = f"""
    pandoc -s --self-contained -t html5
    """.strip()

    if css_path:
        command += f" -c {css_path.absolute()}"

    if _is_existent_path(markdown):
        markdown = Path(markdown).read_text(encoding="utf8")

    result = subprocess.run(
        command,
        universal_newlines=True,
        input=markdown,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        encoding="utf8",
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr)
    html = result.stdout
    html = premailer.transform(html)

    if output_path:
        Path(output_path).write_text(html, encoding="utf8")
    return html


class _CSSCache:
    """CSSCache Folder."""

    def __init__(self, folder=None):
        if folder is None:
            folder = _default_css_cache
        self.folder = Path(folder)
        assert self.folder.exists()

    def __call__(self, arg) -> Union[None, Path]:
        if arg is None:
            return None
        arg = Path(arg)
        if arg.exists():
            return arg
        if (self.folder / arg).exists():
            return self.folder / arg
        candidates = list(self.folder.glob(f"{arg}*"))
        if len(candidates) == 1: 
            return candidates[0]
        elif 1 < len(candidates):
            import warnings

            warnings.warn(
                f"Multiple candidates are detected from ``{arg}``, but one is retured."
            )
            return candidates[0]
        raise ValueError(f"Cannot handle appripriate css file by `{arg}`")


if __name__ == "__main__":
    result = to_html("sample.html", "sample", output_path="output.html")
    print(result)
