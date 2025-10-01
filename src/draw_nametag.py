from collections import namedtuple
from pptx.slide import Slide
from pptx.shapes.shapetree import GroupShapes

from .draw_shape import ShapeDrawer, TextBoxDrawer, AutoShapeDrawer

BoundingBox = namedtuple("BoundingBox", ["left", "top", "width", "height"])

class NameTagDrawer(ShapeDrawer):
    def __init__(self):
        self.drawers: list[ShapeDrawer] = []
        self._bbox: BoundingBox = None

    @staticmethod
    def create_from_slide(slide: Slide):
        nameTagDrawer = NameTagDrawer()
        nameTagDrawer.drawers = list(nameTagDrawer._create_drawers(slide))
        if len(nameTagDrawer.drawers) == 0:
            raise ValueError("No shape found in the slide")
        nameTagDrawer._bbox = nameTagDrawer.get_position()
        nameTagDrawer.to_relative_position(nameTagDrawer.left, nameTagDrawer.top)
        return nameTagDrawer

    @property
    def left(self):
        return self._bbox.left
    
    @property
    def top(self):
        return self._bbox.top
    
    @property
    def width(self):
        return self._bbox.width
    
    @property
    def height(self):
        return self._bbox.height

    def _create_drawers(self, slide: Slide):
        shapes = list(slide.shapes)
        for shape in shapes:
            if __debug__:
                print("(debug) shape:", shape.name, shape.shape_type, end="\n\n")
            sd = ShapeDrawer.create(shape)
            if sd is None:
                continue
            elif isinstance(sd, GroupShapes):
                shapes += list(sd)
                continue
            yield sd
    
    def add_drawer(self, drawer: ShapeDrawer):
        self.drawers.append(drawer)
        self._bbox = self.get_position()
        self.to_relative_position(self.left, self.top)
    
    def get_position(self) -> BoundingBox:
        bbox = _get_rotated_bounding_box(self.drawers[0].shape)
        left:float = bbox.left
        top:float = bbox.top
        right:float = bbox.left + bbox.width
        buttom:float = bbox.top + bbox.height
        
        for drawer in self.drawers[1:]:
            bbox = _get_rotated_bounding_box(drawer.shape)
            left = min(left, bbox.left)
            top = min(top, bbox.top)
            right = max(right, bbox.left + bbox.width)
            buttom = max(buttom, bbox.top + bbox.height)
        return BoundingBox(left, top, right - left, buttom - top)

    def to_relative_position(self, left: float, top: float):
        for drawer in self.drawers:
            drawer.to_relative_position(left, top)

    def draw(self, slide: Slide, left: float=0, top: float=0):
        for drawer in self.drawers:
            drawer.draw(slide, left, top)
    
    def set_text(self, data: dict[str, int|str]):
        for drawer in self.drawers:
            if isinstance(drawer, TextBoxDrawer) or isinstance(drawer, AutoShapeDrawer):
                if drawer.label in data:
                    drawer.set_text(data[drawer.label])

def _get_rotated_bounding_box(shape):
    """
    회전된 shape의 실제 바운딩 박스를 계산합니다.
    
    Args:
        shape: pptx shape 객체
        
    Returns:
        BoundingBox: 실제 바운딩 박스 좌표
    """
    import math
    
    # 원본 shape의 좌표와 크기
    left = shape.left.cm
    top = shape.top.cm
    width = shape.width.cm
    height = shape.height.cm
    rotation = shape.rotation
    
    # 회전이 없으면 원본 좌표 반환
    if rotation == 0:
        return BoundingBox(left, top, width, height)
    
    # 회전 각도를 라디안으로 변환
    angle_rad = math.radians(rotation)
    cos_angle = abs(math.cos(angle_rad))
    sin_angle = abs(math.sin(angle_rad))
    
    # 회전된 바운딩 박스의 크기 계산
    rotated_width = width * cos_angle + height * sin_angle
    rotated_height = width * sin_angle + height * cos_angle
    
    # 회전된 바운딩 박스의 중심점
    center_x = left + width / 2
    center_y = top + height / 2
    
    # 회전된 바운딩 박스의 좌상단 좌표
    rotated_left = center_x - rotated_width / 2
    rotated_top = center_y - rotated_height / 2
    
    return BoundingBox(rotated_left, rotated_top, rotated_width, rotated_height)