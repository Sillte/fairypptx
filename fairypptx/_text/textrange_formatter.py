"""
### Protocol of Formatter.

* Input: __init__(self, object).
* Output: __call__(self, object).
* Must be serialiable by `.pkl`.
"""
from enum import IntEnum
from fairypptx import constants 

SUFFICIENT_SMALL = 0.1


class ParagraphTextRangeFormatter:
    """Following to authors' imagination,    

    -----------------------
    Target: (Font, Paragraphformat).

    Input: Construct the dict of dict
        Key: paragraphtype -> (indentlevel, lineindex) 
        Value: (Font, ParagraphFormat)
    Output: 
        Get the nearst key and  set the format.
    """
    def __init__(self, tr):
        self.data = dict()
        for line_index, para in enumerate(tr.paragraphs):
            p_type = _to_paragraph_type(para)
            key = (para.api.IndentLevel, line_index)
            if p_type not in self.data:
                self.data[p_type] = dict()
            self.data[p_type][key] = (para.font, para.paragraphformat)
        if not self.data:
            raise ValueError("Empty format is not accepted.")

    def __call__(self, tr):
        for line_index, para in enumerate(tr.paragraphs):
            indentlevel = para.api.IndentLevel
            p_type = _to_paragraph_type(para)
            target_key = (para.api.IndentLevel, line_index)
            if p_type not in self.data:
                sub_data = self.data[min(self.data.keys())]
            else:
                sub_data = self.data[p_type]
            dist_key = lambda k: [abs(arg1 - arg2 + SUFFICIENT_SMALL) for arg1, arg2 in zip(k, target_key)]
            key = min(sub_data.keys(), key=dist_key)
            ref_font, ref_paragraphformat = sub_data[key]
            para.font = ref_font
            para.paragraphformat = ref_paragraphformat
        return self


class _ParagraphType(IntEnum):
    BulletNone = 0
    BulletUnNumbered = 1
    BulletNumbered = 2
    BulletPicture = 3


def _to_paragraph_type(para):
    if para.paragraphformat["Bullet.Visible"] == constants.msoTrue:
        if para.paragraphformat["Bullet.Type"] == constants.ppBulletNumbered:
            return _ParagraphType.BulletNumbered
        elif para.paragraphformat["Bullet.Type"] == constants.ppBulletUnnumbered:
            return _ParagraphType.BulletUnNumbered
        elif para.paragraphformat["Bullet.Type"] == constants.ppBulletPicture:
            return _ParagraphType.BulletPicture
    return _ParagraphType.BulletNone
        


if __name__ == "__main__":
    pass
