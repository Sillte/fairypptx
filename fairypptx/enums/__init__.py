from enum import IntEnum

class MsoFillType(IntEnum):
    FillGradient = 3
    FillPatterned = 2
    FillSolid = 1
    FillTextured = 4
    FillBackground = 5
    FillPicture = 6
    FillMixed = -2

class MsoShapeType(IntEnum):
    AutoShape = 1
    Group = 6
    Line = 9
    Picture = 13
    Table = 19
    TextBox = 17
    PlaceHolder = 14
