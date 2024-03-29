""" Handling conversion between `HTML` and `Clipboard`.  

Mainly there 2 operations, `pull` and `push`.    

```python

# Over-writing of Clipboard.
push(html)  # Copy `html` to clipboard.

# Get the content of Clipboard.
html = pull()  # 
```

Reference
---------
* https://gist.github.com/Erreinion/6691093/revisions
* https://docs.microsoft.com/en-us/windows/win32/dataxchg/html-clipboard-format   
"""

from pathlib import Path
import re
import random
import time

import win32clipboard


def push(arg, *, is_path=None):
    """Push the html to Clipboard.
    Args:
        arg: Path or str as  html.
        is_path: Specify how to deal with `arg`.
                 If `True`, arg is regarded as `Path`.
                 If `False`, arg is regarded as `str`.
                 If `None`, inferred, `arg` is path or content.
    """

    if is_path is None:
        html = _to_content(arg)
    elif is_path is True:
        html = Path(arg).read_text(encoding="utf8")
    else:
        html = arg

    if isinstance(arg, Path):
        source = arg
    else:
        source = __file__

    data = _ClipboardHTML(html, source=source)
    data.to_clipboard()


def push_fragment(fragment: str, html: str = None, url=None):
    """Push the part of HTML to `Clipboard`.
    Args:
        fragment: the part of html.
        html: the source of html.
        url: the url of the html address.
    """
    if html is None:
        html = _gen_default_html(fragment)
    if url is None:
        url = Path(__file__).name
    start_index = html.find(fragment)
    end_index = start_index + len(fragment)
    data = _ClipboardHTML(html, start_index, end_index, source=url)
    data.to_clipboard()


def pull():
    """Pull the content's of clipboard."""
    data = _ClipboardHTML.from_clipboard()
    return data.fragment.strip()


#  Register Clipboard Format.
CF_HTML = win32clipboard.RegisterClipboardFormat("HTML Format")

# Control Characters,
CONTROL_CHARS = "".join(chr(elem) for elem in range(32)) + chr(127)


def _gen_default_html(fragment):
    return f"<HTML><HEAD></HEAD><BODY><!--StartFragment-->{fragment}<!--EndFragment--></BODY></HTML>"


def _to_content(arg):
    """Assure the return of `content`.
    If `arg` is regarded as existent,
    `filepath`, then
    its content is returned (UTF8 decoded).
    """
    try:
        path = Path(arg)
        if path.exists():
            return path.read_text(encoding="utf8")
    except OSError:
        pass
    return arg


class _ClipboardHTML:
    """It is a data structure to handle the Clipboard HTML data.


    CAUTION
    -------
    `start_index` and `end_index` are NOT index of `str`, but index of encoded bytes.

    Note
    -------------------------------
    Currently, `Selection` is not handled.
    """

    def __init__(
        self, html, start_index=None, end_index=None, source=None, version=None
    ):
        if isinstance(html, str):
            self.en_html = html.encode("utf8")
        elif isinstance(html, bytes):  # utf8 is assumed.
            self.en_html = html
        else:
            raise TypeError(r"`html` must be `str` or utf-8 encoded `bytes`.")
        if start_index is None:
            start_index = 0
        self.start_index = start_index
        if end_index is None:
            end_index = len(self.en_html)
        self.end_index = end_index
        if source is None:
            source = "_ClipboardHTML"
        self.source = source
        if version is None:
            version = "1.0"
        self.version = version

    def __str__(self):
        d = dict()
        for key in dir(self):
            if key.startswith("_"):
                continue
            d[key] = getattr(self, key)
        return str(d)

    @property
    def html(self):
        return self.en_html.decode("utf8")

    @property
    def fragment(self):
        """Return the target fragment of html."""

        # Maybe `strip` of CONTROL_CHARS is unnecesarry.
        return (
            self.en_html[self.start_index : self.end_index]
            .decode("utf8")
            .strip(CONTROL_CHARS)
        )

    @classmethod
    def from_fragment(cls, fragment, html=None, url=None):
        if html is None:
            html = _gen_default_html(fragment)
        if url is None:
            url = Path(__file__).name
        en_html = html.encode()
        en_frament = html.encode()
        start_index = en_html.find(en_fragment)
        end_index = start_index + len(en_fragment)
        return _ClipboardHTML(en_html, start_index, end_index, source=url)

    @classmethod
    def from_clipboard(cls):
        src = _get_clipboard_data(CF_HTML).decode("utf8")

        # Decompose `description` and `header`.
        s_index = src.find("SourceURL")
        assert s_index != -1
        while src[s_index] != "\n":
            s_index += 1
        description, body = src[: s_index + 1], src[s_index + 1 :]
        assert description + body == src

        lines = [line.strip() for line in description.split("\n") if line]
        header_dict = dict()
        for l, line in enumerate(lines):
            key, value = line.split(":", 1)
            header_dict[key] = value

        kwargs = dict()
        kwargs["html"] = body.encode("utf8")
        s_fragment = int(header_dict["StartHTML"])
        e_fragment = int(header_dict["EndHTML"])
        kwargs["start_index"] = s_fragment - len(description.encode())
        kwargs["end_index"] = e_fragment - len(description.encode())
        kwargs["version"] = header_dict["Version"]
        kwargs["source"] = header_dict["SourceURL"]

        return cls(**kwargs)

    def to_clipboard(self):
        content = self._to_clipboard_content()

        try:
            win32clipboard.OpenClipboard(0)
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(CF_HTML, content.encode("utf8"))
        except Exception as e:
            print("Exception", e)
        finally:
            win32clipboard.CloseClipboard()

    def _to_clipboard_content(self):
        template = (
            "Version:0.9\r\n"
            "StartHTML:%09d\r\n"
            "EndHTML:%09d\r\n"
            "StartFragment:%09d\r\n"
            "EndFragment:%09d\r\n"
            "SourceURL:%s\r\n"
        )
        dummy = template % (0, 0, 0, 0, self.source)
        n_description = len(dummy.encode())
        description = template % (
            n_description,
            n_description + len(self.en_html),
            n_description + self.start_index,
            n_description + self.end_index,
            self.source,
        )
        return description + self.html


def _get_clipboard_data(format_id, max_trial=5):
    """To handle collision of Clipboard Usage.
    Especially, access is denied.
    """
    last_exception = None
    for _ in range(max_trial):
        try:
            win32clipboard.OpenClipboard(0)
            data = win32clipboard.GetClipboardData(format_id)
            return data
        except OSError as e:
            if e.errorno == 5:
                time.sleep(random.random() / 10)
            last_exception = e
        finally:
            win32clipboard.CloseClipboard()
    raise last_exception


# Makeshift test.
def test_basic():
    sample_html = "<p>Writing Fragment to Clipboard</p>"
    push(sample_html, is_path=False)
    output = pull()
    assert sample_html == output

    sample_jp_html = "<p>貴方は潜伏背徳者ですか？ Quoted from Are you a werewolf? </p>"
    push(sample_jp_html, is_path=False)
    output = pull()
    assert sample_jp_html == output


def test_source_jp():
    """
    Case where `source` is multi-byte,
    """
    html = "<p>これでいいのかな..</p>"
    source = "マルチバイトsource"

    data = _ClipboardHTML(html, source=source)
    data.to_clipboard()

    c_data = _ClipboardHTML.from_clipboard()
    assert c_data.fragment == html
    assert c_data.source == source


if __name__ == "__main__":
    test_basic()
    test_source_jp()
