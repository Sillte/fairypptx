""" Here, I'd like to confirm whether it is possible to convert `fairypptx.Markdown` to `Markdown`.  


Reference
------
# https://hackage.haskell.org/package/pandoc-types-1.12.4.5/docs/Text-Pandoc-Definition.html

"""
from itertools import groupby
from pathlib import Path
import json
import subprocess
from enum import Enum, auto
from fairypptx import constants 

ITEMIZATION_TYPES = [constants.ppBulletUnnumbered, constants.ppBulletNumbered]

class Kind:
    """ `Kind` represents the type of `paragraphs`s. 
    """
    NORMAL = auto()
    ITEMIZATION = auto()


class Converter: 
    _pandoc_api = None
    def __init__(self):
        pass

    @classmethod
    def _from_jsonast(cls, jsonast):
        # These metadata seems mandatory... 
        #
        if "meta" not in jsonast:
            jsonast["meta"] = {}
        if "pandoc-api-version" not in jsonast:
            jsonast["pandoc-api-version"] = cls._pick_pandoc_api()
        from pprint import pprint
        print("JSONAST")
        pprint(jsonast["blocks"])
        content = json.dumps(jsonast)
        ret = subprocess.run("pandoc -t gfm -f json",
                              universal_newlines=True, 
                              stdout=subprocess.PIPE, 
                              input=content, encoding="utf8")
        assert ret.returncode == 0
        return ret.stdout


    def parse(self, markdown): 
        shape = markdown.shape
        blocks = []
        textrange = shape.textrange
        queue_paragraphs = to_kind_groups(textrange)  # List[Dict[str, List["paragraphs"]]]
        for kind, paragraphs in queue_paragraphs:
            if kind == Kind.ITEMIZATION:
                blocks += itemization_paragraphs_to_blocks(paragraphs)
            else:
                # Fallback and  `normal` cases. 
                blocks += normal_paragraphs_to_blocks(paragraphs)
        jsonast = {"blocks": blocks}
        return self._from_jsonast(jsonast)

    @classmethod
    def _pick_pandoc_api(cls):
        """Get `pandoc` api version.
        I imagine that more  
        and, if api format changes, 
        then this function corrupts...
        """
        if cls._pandoc_api is not None:
            return cls._pandoc_api
        ret = subprocess.run("pandoc -t json",
                              universal_newlines=True, 
                              stdout=subprocess.PIPE, 
                              input="", encoding="utf8")
        cls._pandoc_api = json.loads(ret.stdout)["pandoc-api-version"]
        return cls._pandoc_api
        

def to_kind_groups(textrange):
    """  
    * Successive paragraphs are treated as the same structures.

    [TODO] If you consider the existence of `HEADER`, 
    you should take into account the change of `Kind`.
    """
    paragraphs = textrange.paragraphs
    def _to_kind(paragraph):
        if paragraph.api.ParagraphFormat.Bullet.Type in ITEMIZATION_TYPES:
            return Kind.ITEMIZATION
        else:
            return Kind.NORMAL

    result = [(key, list(values)) for key, values
               in groupby(paragraphs, key=lambda p: _to_kind(p))]
    return result


def normal_paragraphs_to_blocks(paragraphs):
    """ Normally convert to `paragraphs`. 
    """
    blocks = []
    for paragraph in paragraphs:
        block = {"t": "Para"}
        inlines = sum((run_to_inlines(run) for run in paragraph.runs), [])
        block["c"] = inlines
        blocks.append(block)
    return blocks

def itemization_paragraphs_to_blocks(paragraphs):
    """
    """
    # All of the `paragpraph.ParagraphFormat.Bullet.Type` must belong to `itemization`.
    assert all(para.api.ParagraphFormat.Bullet.Type in ITEMIZATION_TYPES for para in paragraphs)

    # It seems one `blocks` correspond to one element of the specified structure.
    # * L1 ..blocks `A`
    # * L2 .. blocks `B`
    #    * L2-1 .. blocks `B`
    # * L3 .. Blocks `C`

    def _from_blocks(blocks, root_paragraph):
        """Generate one `block` of `BulletList` or `OrderedList`
        from the elements o `blocks`.  
        """
        def _gen_ordered_list_block(blocks):
            root = {"t": "OrderedList"}
            list_attributes = [1, {'t': 'Decimal'}, {'t': 'Period'}]
            root["c"] = [list_attributes, blocks]
            return root

        def _gen_bullest_list_block(blocks):
            root = {"t": "BulletList"}
            root["c"] = blocks
            return root
        if root_paragraph.api.ParagraphFormat.Bullet.Type == constants.ppBulletNumbered:
            return _gen_ordered_list_block(blocks)
        else:
            return _gen_bullest_list_block(blocks)

    def _to_partitions(si, ei):
        """Based on the `ParagraphFormat.Bullet.Type`,
        it divides `[si, ei)`  to  multiple intervals.
        """

        def _to_bullet_type(i):
            return paragraphs[i].api.ParagraphFormat.Bullet.Type
        def _to_indent_level(i):
            return paragraphs[i].api.IndentLevel
        def _is_changed_pivot(i, prev_bullet_type):
            return (prev_bullet_type != _to_bullet_type(i)
                    and target_indent_level == _to_indent_level(i))
        target_indent_level = _to_indent_level(si)

        prev_bullet_type = _to_bullet_type(si)
        prev_pivot = si
        pivot = si + 1
        result = []
        while pivot < ei:
            if _is_changed_pivot(pivot, prev_bullet_type):
                result.append((prev_pivot, pivot))
                prev_bullet_type = _to_bullet_type(pivot)
                prev_pivot = pivot
                pivot = prev_pivot + 1
            else:
                pivot += 1
        result.append((prev_pivot, ei))
        assert result[0][0] == si
        assert result[-1][1] == ei
        return result

    def _to_blocks(si, ei):
        """ Convert to `blocks` `[si, ei)`'s paragraphs to `blocks`.
        """ 

        def _to_element_end(x):
            # Return index of end of `element`.
            y = x + 1
            while y < ei:
                if paragraphs[x].api.IndentLevel >= paragraphs[y].api.IndentLevel:
                    break
                y += 1
            return y
         
        elements = []
        x = si
        while x < ei:
            y = _to_element_end(x)
            if y - x == 1:
                target = {"t": "Plain"}
                target["c"] = sum((run_to_inlines(run) for run in paragraphs[x].runs), [])
                elements.append([target])
            else:
                block = {"t": "Plain"}
                block["c"] = sum((run_to_inlines(run) for run in paragraphs[x].runs), [])
                blocks = [block]

                # For example, in these situations,  
                # multiple `blocks` are required. 
                # `_to_partitions` are used to separate `Unnumber` and `numbers`
                # keeping in mind the `IndentLevel`.  
                # * ITEM1 
                #   * ITEM1-1
                #   - ITEM1-2
                for rs, re in _to_partitions(x + 1, y):
                    inner_blocks = _to_blocks(rs, re)
                    target = _from_blocks(inner_blocks, paragraphs[rs])
                    blocks.append(target)
                elements.append(blocks)
            x = y
        return elements

    blocks = _to_blocks(0, len(paragraphs))
    root = _from_blocks(blocks, paragraphs[0])
    return [root]


def paragraph_to_block(paragraph):
    data = dict()
    data["t"] = "Para"
    inlines = []
    for run in paragraph.runs:
        inlines += run_to_inlines(run)
    data["c"] = inlines
    return data



def run_to_inlines(textrange):
    """ Here, only focuses on the style of `string`. 
    """
    def _to_str(s):
        return {"t": "Str", "c": str(s)}

    text = str(textrange.api.Text).rstrip("\r")
    tokens = text.split("\013")  # vertical tag.
    result = []
    for i, token in enumerate(tokens):
        result += [_to_str(token)]
        if i != len(tokens) - 1:
            result += [{"t": "LineBreak"}]

    if textrange.font.bold:
        result = [{"t": "Strong", "c": result}]

    return result
    

def to_script(markdown: "Markdown"):
    """Public API.
    """
    return Converter().parse(markdown)


if __name__ == "__main__":
    from fairypptx import Markdown

    
    TEXT = """
* NUM1
    * ITEM1
    * ITEM2
        1. ITEM2
        2. ITEM2
""".strip()
    data = dict()
    data["blocks"] = [{'t': 'Para', 'c': [{"t": "Str", "c": "sample"}]}]
    #ret = Converter._from_jsonast(data)
    #print("ret", ret)

    markdown = Markdown.make(TEXT)
    ret = to_script(markdown)
    
    print("result----")
    from pprint import pprint
    print(ret)
    print("result----")
    Path("output.md").write_text(ret)
    exit(0)

    textrange = markdown.shape.textrange
    for p in textrange.paragraphs:
        for u in p.runs:
            print(u.text)
#        print(p.text)
        

