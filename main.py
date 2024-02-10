from io import BytesIO
from os import path
from openpyxl import load_workbook
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

import settable
from analyze_slide import get_slide_info
from utils import chunk_list, tuples_to_dict_list



class NameTag:
    def __init__(self, pptx):
        self._filename = path.basename(pptx)
        self.prs = Presentation(pptx)
        self.samples = []
        for slide in self.prs.slides:
            self.samples.append(get_slide_info(slide))
        self.sample_info = self.samples[0]
        self.num_col, self.num_row = self.get_max_col_row()
        self.num_per_slide = self.num_col * self.num_row
        self.left, self.top = self.get_start_pos()

    def get_max_col_row(self):
        num_col = self.prs.slide_width.cm // self.sample_info.width
        num_row = self.prs.slide_height.cm // self.sample_info.height
        return (int(num_col), int(num_row))

    def get_start_pos(self):
        left = (self.prs.slide_width.cm - self.sample_info.width * self.num_col) / 2
        top = (self.prs.slide_height.cm - self.sample_info.height * self.num_row) / 2
        return (left, top)

    def add_nametag_info(self, slide, data):
        for i, d in enumerate(data):
            if i == self.num_per_slide:
                break
            col_idx = i % 2
            row_idx = i // 2
            
            nametag_form = self.samples[d["sample num"]]
            for image_form in nametag_form.images:
                slide.shapes.add_picture(
                    BytesIO(image_form.image.blob),
                    Cm(self.left + col_idx * nametag_form.width + image_form.left),
                    Cm(self.top + row_idx * nametag_form.height + image_form.top),
                    Cm(image_form.width),
                    Cm(image_form.height))

            for text_form in nametag_form.texts:
                p = slide.shapes.add_textbox(
                    Cm(self.left + col_idx * nametag_form.width + text_form.left),
                    Cm(self.top + row_idx * nametag_form.height + text_form.top),
                    Cm(text_form.width),
                    Cm(text_form.height)
                ).text_frame.paragraphs[0]
                p.alignment = text_form.alignment
                r = p.add_run()
                r.font = text_form.font
                r.text = d[text_form.text.lower()]
                
    def save(self):
        self.prs.save(f'dist/generated-{self._filename}') 

def read_excel_data(filename):
    workbook = load_workbook(filename, data_only=True)
    sheet = workbook.active
    data = []
    for row in sheet.iter_rows(values_only=True):
        data.append(row)

    for i, row in enumerate(data):
        data[i] = (c if c is not None else "" for c in row)

    header = data[0]
    header = [h.lower() for h in header]

        
    return header, data[1:]

excel_filename = 'example/attendees_list-example.xlsx'  # 엑셀 파일명 입력
header, nametag_data = read_excel_data(excel_filename)

try:
    sample_num_idx = header.index("sample num")
except:
    header += ("sample num",)
    nametag_data = [d + (0,) for d in nametag_data]
else:
    for i, d in enumerate(nametag_data):
        d = list(d)
        if d[sample_num_idx]:
            d[sample_num_idx] = int(d[sample_num_idx])
        else: 
            d[sample_num_idx] = 0
        nametag_data[i] = tuple(d)

nametag_data = tuples_to_dict_list(header, nametag_data)

pptx_filename = 'example/nametag-example.pptx' 
nametag = NameTag(pptx_filename)
nametag_data = chunk_list(nametag_data, nametag.num_per_slide)

slide_layout = nametag.prs.slide_layouts[0]
for d in nametag_data:
    slide = nametag.prs.slides.add_slide(slide_layout)
    nametag.add_nametag_info(slide, d)

nametag.save()
