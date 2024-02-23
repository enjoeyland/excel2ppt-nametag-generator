import os

from pptx import Presentation
from tkinter import filedialog

from src.draw import NameTagDrawer
from src.utils import get_data_by_sample, open_file_with_default_program

import argparse

def get_args():
    parser = argparse.ArgumentParser(description='Create nametag pptx from excel file')
    parser.add_argument('-excel', required=True, help='Excel file name')
    parser.add_argument('-pptx', required=True, help='PowerPoint file name')
    return parser.parse_args()

if __name__ == "__main__":
    args = get_args()

    data_by_sample = get_data_by_sample(args.excel)

    prs = Presentation(args.pptx)

    for i in data_by_sample.keys():
        NameTagDrawer(prs, i, data_by_sample[i]).draw()
    
    filename = filedialog.asksaveasfilename(
        defaultextension=".pptx",
        filetypes=[("PowerPoint files", "*.pptx")],
        title="Save the file as",
        initialfile=f"generated-{os.path.basename(args.pptx)}"
    )
    if filename:
        prs.save(filename)
        print(f"Done: '{filename}' is saved")
        open_file_with_default_program(filename)
    else:
        print("Canceled by user.")