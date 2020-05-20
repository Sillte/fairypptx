"""Handling processing related to location and position. 

"""

from fairypptx.box import Box

class LocationAdjuster:
    """ Focusing on shape, 
    determine the size of shape. 
    """
    def __init__(self, shape):
        self.shape = shape

    def center(self):
        target_width = self.shape.box.width
        target_height = self.shape.box.height
        slide = self.shape.slide
        slide_width = slide.box.width
        slide_height = slide.box.height
        left = (slide_width - target_width) / 2
        top = (slide_height - target_height) / 2
        self.shape.api.Left = left
        self.shape.api.Top = top
        return self.shape
