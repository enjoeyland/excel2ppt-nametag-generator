from os import path

from pptx import Presentation

from src.draw import NameTagDrawer
from src.utils import get_data_by_sample

if __name__ == "__main__":
    excel_filename = 'example/attendees_list-example.xlsx'  # 엑셀 파일명 입력
    data_by_sample = get_data_by_sample(excel_filename)

    pptx = 'example/nametag-example.pptx' 
    prs = Presentation(pptx)

    for i in data_by_sample.keys():
        NameTagDrawer(prs, i, data_by_sample[i]).draw()
    
    filename = path.basename(pptx)
    prs.save(f'dist/generated-{filename}') 