from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.enum.text import PP_ALIGN

from src.utils import dotdict

def get_image_info(shape):
    if shape.is_placeholder and  shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
        return dotdict({
            "type": "image", 
            "left": shape.left.cm,
            "top": shape.top.cm,
            "width" : shape.width.cm,
            "height": shape.height.cm,
            "image": shape.image
        })

def get_textbox_info(shape):
    if shape.has_text_frame:
        return dotdict({
            "type": "text", 
            "left": shape.left.cm,
            "top": shape.top.cm,
            "width" : shape.width.cm,
            "height": shape.height.cm,
            "text": shape.text_frame.text,
            "alignment": shape.text_frame.paragraphs[0].alignment or PP_ALIGN.LEFT,
            "font": shape.text_frame.paragraphs[0].runs[0].font
        })

def to_relative_info(absolute_info):
    result = dotdict({
        "type": "layout",
        "images": [],
        "texts": []
    })
    basis =  absolute_info[0].copy()
    for info in absolute_info:
        if info.left < basis.left and info.top < basis.top:
            basis =  info.copy()

    right = basis.left + basis.width
    buttom = basis.top + basis.height
    for info in absolute_info:
        right = max(right, info.left + info.width)
        buttom = max(buttom, info.top + info.height)

        info.left -= basis.left
        info.top -= basis.top
        if info.type == "text":
            result.texts.append(info)
        elif info.type == "image":
            result.images.append(info)
    
    result.width = right - basis.left
    result.height = buttom - basis.top
    return result

def get_slide_info(slide):
    absolute_info = []
    for shape in slide.shapes:
        # print("shape:", shape.name, shape.shape_type)
        absolute_info.append(get_image_info(shape) or get_textbox_info(shape))
    return to_relative_info(absolute_info)

if __name__ == "__main__":
    pptx_filename = 'example/nametag-example.pptx' 
    prs = Presentation(pptx_filename)
    
    for i, slide in enumerate(prs.slides):
        print(f"{'='*5}#{i} slide{'='*5}")
        print(get_slide_info(slide))
