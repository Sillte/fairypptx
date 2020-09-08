"""Test for Markdown handling. 

(2020-04-26) Currently, both of library and test are insufficient.   
(2020-09-08) A little progress. 
"""
import pytest
from fairypptx import Markdown
from fairypptx import Shapes


def test_make():
    text = """## Sample Header.  
This is a sample documentation.

|ITEM1|ITEM2|
|-----|------|
|Key1|Key2|

This case the return is not Markdown, but shapes,   
because `Table` and `Texts` is separate. 
* [link](https://ruru-jinro.net/)
""".strip()
    shapes = Markdown.make(text)
    assert isinstance(shapes, Shapes) 
    assert len(shapes) == 2, "Table and `Normal`. "

    text = """## Sample Header.  
This case the return is  Markdown, since
only Table exists.
"""
    markdown = Markdown.make(text)
    assert isinstance(markdown, Markdown) 

def test_interpret():
    # Test of interpret.
    text = """
### sample   

* ITEM1  
    - ITEM1-1  
* ITEM2  
    """.strip()
    str(Markdown.make(text))

if __name__ == "__main__":
    pytest.main([__file__, "--capture=no"])


