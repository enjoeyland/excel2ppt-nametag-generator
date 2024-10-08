import os
import argparse
import tkinter as tk

from pptx import Presentation
from tkinter import filedialog

from src.draw_slide import SlideDrawer
from src.utils import get_data_by_sample, open_file_with_default_program
from src.gui import get_args_by_gui


def get_args():
    parser = argparse.ArgumentParser(description='Create nametag pptx from excel file')
    parser.add_argument('--excel', type=str, help='Excel file name')
    parser.add_argument('--pptx', type=str, help='PowerPoint file name')
    parser.add_argument('--margin_x', type=float, default=0.0, help='Margin between nametags in x direction. unit: cm')
    parser.add_argument('--margin_y', type=float, default=0.0, help='Margin between nametags in y direction. unit: cm')
    parser.add_argument('--padding_x', type=float, default=0.0, help='Padding of nametag in x direction. unit: cm')
    parser.add_argument('--padding_y', type=float, default=0.0, help='Padding of nametag in y direction. unit: cm')
    parser.add_argument('--per_slide', type=int, help='Number of nametags per slide')
    parser.add_argument('--gui', action='store_true', help='Use GUI to select files and set parameters')
    args = parser.parse_args()

    if not args.excel or not args.pptx or args.gui:
        args = get_args_by_gui(args)
        args.gui = True

    return args

if __name__ == "__main__":
    args = get_args()

    data_by_sample = get_data_by_sample(args.excel)

    prs = Presentation(args.pptx)

    sample_num = len(prs.slides)
    for i in data_by_sample.keys():
        if isinstance(i, str):
            print(f"Warning: Sample number '{i}' is not an integer. Skip drawing for sample '{i}'.")
            continue
        if i >= sample_num:
            print(f"Warning: No sample slide with index {i} exists. Skip drawing for sample {i}.")
            print("Hint: value of cloumn 'Sample Num' should start from 0 and be continuous.")
            if args.gui:
               tk.messagebox.showwarning("Warning", f"Warning: No sample slide with index {i} exists. Skip drawing for sample {i}.\nHint: value of cloumn 'Sample Num' should start from 0 and be continuous.")
            continue
        SlideDrawer(prs, i, data_by_sample[i]).draw(
            margin=(args.margin_x, args.margin_y),
            padding=(args.padding_x, args.padding_y),
            per_slide=args.per_slide
        )
    
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