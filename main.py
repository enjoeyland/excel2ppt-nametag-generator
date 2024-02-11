from io import BytesIO
import math
from os import path
from collections import defaultdict

from openpyxl import load_workbook
from pptx import Presentation
from pptx.util import Cm

import src.settable
from src.analyze_slide import get_slide_info
from src.utils import chunk_list, tuples_to_dict_list


class Position:
    def __init__(self, slide_size, sample, data):
        self._slide_width, self._slide_height = slide_size
        self._sample = sample

        self.num_col, self.num_row = self.get_max_col_row()
        self.num_per_slide = self.num_col * self.num_row
        self.left, self.top = self._get_start_pos()
        self.data_by_slide = chunk_list(data, self.num_per_slide)


    def get_max_col_row(self):
        num_col = self._slide_width / self._sample.width
        num_row = self._slide_height / self._sample.height
        return (int(num_col), int(num_row))

    def _get_start_pos(self):
        left = (self._slide_width - self.num_col * self._sample.width) / 2
        top = (self._slide_height - self.num_row * self._sample.height) / 2
        return (left, top)
    
    def _get_index(self, idx):
        col_idx = idx % self.num_col
        row_idx = idx // self.num_col
        return (col_idx, row_idx)
    
    def _get_position(self, idx):
        col_idx, row_idx = self._get_index(idx)
        return (self.left + col_idx * self._sample.width, self.top + row_idx * self._sample.height)

    def slide_info_generator(self):
        for data in self.data_by_slide:
            yield self._nametag_info_generator(data)

    def _nametag_info_generator(self, slide_info):
        for idx, d in enumerate(slide_info):
            yield *self._get_position(idx), d

class NameTagGenerator:
    def __init__(self, prs, sample_num, data):
        self._prs = prs
        self.data = data
        try:
            self.sample = get_slide_info(self._prs.slides[sample_num])
        except:
            raise ValueError(f"No sample slide with index {sample_num} exists")

    def add_nametag_info(self, slide, slide_info):
        for left, top, data in slide_info:
            for image_form in self.sample.images:
                slide.shapes.add_picture(
                    BytesIO(image_form.image.blob),
                    Cm(left + image_form.left),
                    Cm(top + image_form.top),
                    Cm(image_form.width),
                    Cm(image_form.height))

            for text_form in self.sample.texts:
                p = slide.shapes.add_textbox(
                    Cm(left + text_form.left),
                    Cm(top + text_form.top),
                    Cm(text_form.width),
                    Cm(text_form.height)
                ).text_frame.paragraphs[0]
                p.alignment = text_form.alignment
                r = p.add_run()
                r.font = text_form.font
                r.text = data[text_form.text.lower()]

    def start(self, blank_slide_layout = 0):
        self.position = Position((self._prs.slide_width.cm, self._prs.slide_height.cm), self.sample, self.data)
        slide_layout = self._prs.slide_layouts[blank_slide_layout]

        for slide_info in self.position.slide_info_generator():
            slide = self._prs.slides.add_slide(slide_layout)
            self.add_nametag_info(slide, slide_info)
        return self._prs

def read_excel_data(filename):
    workbook = load_workbook(filename, data_only=True)
    sheet = workbook.active
    data = []
    for row in sheet.iter_rows(values_only=True):
        data.append(row)

    for i, row in enumerate(data):
        data[i] = tuple(c if c is not None else "" for c in row)

    header = data[0]
    header = [h.lower() for h in header]
        
    return header, data[1:]

def headed_data_with_sample_num(header, data):
    try:
        sample_num_idx = header.index("sample num")
    except:
        header += ("sample num",)
        data = [d + (0,) for d in data]
    else:
        for i, d in enumerate(data):
            d = list(d)
            if d[sample_num_idx]:
                d[sample_num_idx] = int(d[sample_num_idx])
            else: 
                data[i] = tuple(d)
    data = tuples_to_dict_list(header, data)
    return data

def group_by_sample(data):
    data_by_sample = defaultdict(list)
    for d in data:
        data_by_sample[d["sample num"]].append(d)
    return data_by_sample

if __name__ == "__main__":
    excel_filename = 'example/attendees_list-example.xlsx'  # 엑셀 파일명 입력
    header, row_data = read_excel_data(excel_filename)
    headed_data = headed_data_with_sample_num(header, row_data)
    data_by_sample = group_by_sample(headed_data)

    pptx = 'example/nametag-example.pptx' 
    prs = Presentation(pptx)

    for i in data_by_sample.keys():
        generator = NameTagGenerator(prs, i, data_by_sample[i])
        generator.start()
    
    filename = path.basename(pptx)
    prs.save(f'dist/generated-{filename}') 