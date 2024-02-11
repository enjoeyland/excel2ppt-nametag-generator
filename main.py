import os

from pptx import Presentation

from src.draw import NameTagDrawer
from src.utils import get_data_by_sample

import argparse

def get_args():
    parser = argparse.ArgumentParser(description='Create nametag pptx from excel file')
    parser.add_argument('excel', help='Excel file name')
    parser.add_argument('pptx', help='PPTX file name')
    return parser.parse_args()

if __name__ == "__main__":
    args = get_args()

    data_by_sample = get_data_by_sample(args.excel)

    prs = Presentation(args.pptx)

    for i in data_by_sample.keys():
        NameTagDrawer(prs, i, data_by_sample[i]).draw()
    
    filename = os.path.basename(args.pptx)
    if not os.path.exists('dist'):
        os.makedirs('dist')
    prs.save(f'dist/generated-{filename}')
    print(f"generated-{filename} is saved in dist folder")