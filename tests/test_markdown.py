"""Test for Markdown handling. 

(2020-04-26) Currently, both of library and test are insufficient.   
"""
import pytest
from fairypptx import Markdown


def test_make_interpret():
    text = """## Sample Header.  
    This is a sample documentation.
    """.strip()
    shape = Markdown.make(text)
    s = str(Markdown(shape))

if __name__ == "__main__":
    pytest.main([__file__, "--capture=no"])
    pass



