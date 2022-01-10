"""
Notice that the conversion between `fairypptx.Markdown` and `text of Markdown`.
"""

import pytest
import numpy as np
import json
import subprocess
from fairypptx.parts.markdown import Markdown

def _to_jsonast(markdown_script):
    ret = subprocess.run("pandoc -t json",
                          universal_newlines=True, 
                          stdout=subprocess.PIPE, 
                          input=markdown_script, encoding="utf8")
    assert ret.returncode == 0
    return json.loads(ret.stdout)


def test_naive():
    IN_SCRIPT = """ 
Hello world!
    """.strip()
    markdown = Markdown.make(IN_SCRIPT)
    out_script = str(markdown.script)
    assert _to_jsonast(IN_SCRIPT) == _to_jsonast(out_script)

def test_simple_itemization():
    IN_SCRIPT = """ 
* ITEM1
* ITEM2
* ITEM3
""".strip()
    markdown = Markdown.make(IN_SCRIPT)
    out_script = str(markdown.script)
    assert _to_jsonast(IN_SCRIPT) == _to_jsonast(out_script)


if __name__ == "__main__":
    pytest.main([__file__, "--capture=no"])
