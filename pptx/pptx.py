from pptx.font import FontManager

from pptx.group import GroupManager

from pptx.image import ImageManager

from pptx.shape import ShapeManager

from pptx.slide import SlideManager

from pptx.table import TableManager


class PPTXManager(
    FontManager,
    GroupManager,
    ImageManager,
    ShapeManager,
    SlideManager,
    TableManager,):

    def __init__(self):
        super(PPTXManager, self).__init__()