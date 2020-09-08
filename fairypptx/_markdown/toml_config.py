"""Handles conversion of `markdown` 

Here, `toml` is used to communicate `attributes` as for `markdown`.  
Main purpose is to describe `style` of the documentation.  
  

* `toml`'s data is returned as Python primitive types.

Speficiation
-------------
If the firstline of the document starts with ```,  
and it is a valid toml, then these are regarded as  
`attributes` of the document. 

Requirement
-------------
* `pandoc`.

Example

```toml
key1 = "value"
```

This is a sample documentation.
"""

import re
import toml
import itertools


def unite(markdown, config):
    """Integrate `markdown` and `toml`.
    """
    raise RuntimeError()


_pattern = re.compile(r"^```([\w|\s]*)\n((?:^(?!```).*\n)*)```\s*", re.MULTILINE)


def separate(text):
    """Separate `markdown` and `toml`.
    """
    m = _pattern.match(text)
    if not m:
        markdown, config = text, dict()
        return markdown, config
    filetype, filecontent = m.groups()
    print(filetype)
    print(filecontent)
    config = toml.loads(filecontent)
    # toml = text[m.start():m.end()]
    markdown = text[m.end():].strip("\n").strip(" ")
    # print("mark", markdown)
    # print("toml", toml)
    return markdown, config


TEXT = """```toml
key = "label"
value = "fafa"
[[table]]
name = "hoge"
[[table]]
name = "hoge2"
```
fafaefad

```
documentc
```

"""

if __name__ == "__main__":
    separate(TEXT)
