from pptx.slide import Slide

from src.draw_shape import ShapeDrawer, TextBoxDrawer, AutoShapeDrawer

class NameTagDrawer(ShapeDrawer):
    def __init__(self):
        self.drawers: list[ShapeDrawer] = []

    @staticmethod
    def create_from_slide(slide: Slide):
        nameTagDrawer = NameTagDrawer()
        nameTagDrawer.drawers = list(nameTagDrawer._create_drawers(slide))
        nameTagDrawer.left, nameTagDrawer.top, nameTagDrawer.width, nameTagDrawer.height = nameTagDrawer.get_position()
        nameTagDrawer.to_relative_position(nameTagDrawer.left, nameTagDrawer.top)
        return nameTagDrawer

    def _create_drawers(self, slide: Slide):
        shapes = slide.shapes
        for shape in shapes:
            print("shape:", shape.name, shape.shape_type)
            yield ShapeDrawer.create(shape)
    
    def add_drawer(self, drawer: ShapeDrawer):
        self.drawers.append(drawer)
        self.left, self.top, self.width, self.height = self.get_position()
        self.to_relative_position(self.left, self.top)
    
    def get_position(self):
        left:float = self.drawers[0].shape.left.cm
        top:float = self.drawers[0].shape.top.cm
        right:float = self.drawers[0].shape.left.cm + self.drawers[0].shape.width.cm
        buttom:float = self.drawers[0].shape.top.cm + self.drawers[0].shape.height.cm
        for drawer in self.drawers:
            left = min(left, drawer.shape.left.cm)
            top = min(top, drawer.shape.top.cm)
            right = max(right, drawer.shape.left.cm + drawer.shape.width.cm)
            buttom = max(buttom, drawer.shape.top.cm + drawer.shape.height.cm)
        return left, top, right - left, buttom - top

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