from openpyxl import load_workbook
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

def read_excel_data(filename):
    workbook = load_workbook(filename)
    sheet = workbook.active
    data = []
    for row in sheet.iter_rows(values_only=True):
        data.append(row)
    return data[0], data[1:]


def add_badge_info(slide, nametag_info, size):
    for i, shape in enumerate(slide.placeholders):
        placeholder = slide.placeholders[shape.placeholder_format.idx]
        picture = placeholder.insert_picture(nametag_info[i % len(nametag_info)][3])
        
        if size:
            new_width, new_height = size
            picture.left, picture.top = (picture.left + (1 - i % 2) * (picture.width - new_width), picture.top + (1 - i // 2) * (picture.height - new_height))
            picture.width = new_width
            picture.height = new_height

    left = Cm(0.525)
    top = Cm(1.4)
    col_width, row_height = size
    
    for i, (캠퍼스, 이름, 순장, _, color) in enumerate(nametag_info):
        col_idx = i % 2
        row_idx = i // 2
        
        left_offset = col_idx * col_width
        top_offset = row_idx * row_height 
        
        text_left_offset = 0
        text_top_offset = Cm(4.55) - top
        textbox = slide.shapes.add_textbox(left + left_offset + text_left_offset, top + top_offset + text_top_offset, col_width, Pt(40))
        text_frame = textbox.text_frame
        p = text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run_이름 = p.add_run()
        run_이름.text = 이름
        run_이름.font.name = '조선신명조'
        run_이름.font.size = Pt(32)

        text_left_offset = Cm(0.91) - left
        text_top_offset = Cm(11.86) - top
        textbox = slide.shapes.add_textbox(left + left_offset + text_left_offset, top + top_offset + text_top_offset, col_width, Pt(18))
        text_frame = textbox.text_frame
        p = text_frame.paragraphs[0]
        run_캠퍼스_순장 = p.add_run()
        run_캠퍼스_순장.text = f'{캠퍼스} {순장}'
        run_캠퍼스_순장.font.name = '조선신명조'
        run_캠퍼스_순장.font.size = Pt(16)
        run_캠퍼스_순장.font.color.rgb = color
        

def main():
    # 엑셀 파일에서 데이터 읽기
    excel_filename = '순리캠 명단.xlsx'  # 엑셀 파일명 입력
    header, nametag_data = read_excel_data(excel_filename)

    image_path = {
        '간사': '순리캠 명찰_간사.png',
        '순장': '순리캠 명찰_순장.png',
        '예비순장': '순리캠 명찰_예비순장.png',
    }
    main_color = {
        '간사': RGBColor(22, 46, 119),
        '순장': RGBColor(25, 69, 37),
        '예비순장': RGBColor(196, 63, 29),
    }
    nametag_data = [(*nt, image_path[nt[2]], main_color[nt[2]]) for nt in nametag_data]

    pptx_filename = '순리캠 명찰.pptx' 
    prs = Presentation(pptx_filename)


    nametag_size = (Cm(9.0), Cm(11.8))

    num_columns = 2
    num_rows = 2
    nametag_per_slide = num_columns * num_rows
    for i in range(0, len(nametag_data), nametag_per_slide):
        nametag_info = nametag_data[i:i+nametag_per_slide]

        slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(slide_layout)
        
        add_badge_info(slide, nametag_info, nametag_size)
    
    prs.save('생성된 순리캠 명찰.pptx')  # 출력 파일명 입력

if __name__ == "__main__":
    main()
