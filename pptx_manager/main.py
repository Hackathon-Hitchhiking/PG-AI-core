from pptx_manager.font import TextFrameManager
from pptx_manager.group import GroupManager
from pptx_manager.image import ImageManager
from pptx_manager.shape import ShapeManager
from pptx_manager.slide import SlideManager
from pptx_manager.table import TableManager


class PPTXManager(
    TextFrameManager,
    GroupManager,
    ImageManager,
    ShapeManager,
    SlideManager,
    TableManager,
):
    def __init__(self):
        super(PPTXManager, self).__init__()
