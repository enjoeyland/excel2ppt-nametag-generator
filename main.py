import os
import sys
import json
import logging
import argparse

from dataclasses import dataclass
from pptx import Presentation
from tkinter import filedialog, messagebox

from src.draw_slide import SlideDrawer
from src.utils import get_data_by_sample, open_file_with_default_program
from src.gui import get_args_by_gui

logging.basicConfig(level=logging.WARNING, format='%(levelname)s: %(message)s')

def get_args():
    parser = argparse.ArgumentParser(description="Create nametag pptx from excel file")
    parser.add_argument("--excel", type=str, help="Excel file name")
    parser.add_argument("--pptx", type=str, help="PowerPoint file name")
    parser.add_argument("--margin_x", type=float, default=0.0, help="Margin between nametags in x direction. unit: cm")
    parser.add_argument("--margin_y", type=float, default=0.0, help="Margin between nametags in y direction. unit: cm")
    parser.add_argument("--padding_x", type=float, default=0.0, help="Padding of nametag in x direction. unit: cm")
    parser.add_argument("--padding_y", type=float, default=0.0, help="Padding of nametag in y direction. unit: cm")
    parser.add_argument("--per_slide", type=int, help="Number of nametags per slide")
    parser.add_argument("--gui", action="store_true", help="Use Tkinter GUI to select files and set parameters")
    parser.add_argument("--rpc", action="store_true", help="Pass arguments through JSON (Electron IPC)") # RPC: Remote Procedure Call
    args = parser.parse_args()

    if args.rpc:
        args.gui = True
    elif not args.excel or not args.pptx or args.gui:
        args = get_args_by_gui(args)
        args.gui = True
    return args

@dataclass
class GenerateRequest:
    pptx: str
    excel: str
    margin_x: float = 0.0
    margin_y: float = 0.0
    padding_x: float = 0.0
    padding_y: float = 0.0
    per_slide: int = None

class TaskManger:
    def __init__(self, is_gui):
        self.is_gui = is_gui
        self.tasks = {
            "generate_pptx": (self.generate_pptx, GenerateRequest),
        }

    def process_request(self, request):
        task = request.get("task")
        data = request.get("data", {})

        response = {"task": task}
        if task in self.tasks:
            function, dataclass_type = self.tasks[task]
            try:
                request_data = dataclass_type(**data)
            except TypeError as e:
                response.update({"status": "developer_error", "message": f"Invalid parameters for {task}: {str(e)}"})
                response = {"status": "developer_error", "message": f"Invalid parameters for {task}: {str(e)}", "task": task}
            else:
                _response = function(request_data)
                response.update(_response)
        else:
            response.update({"status": "developer_error", "message": f"Unknown task: {task}"})

        print(json.dumps(response))
       
    def generate_pptx(self, data: GenerateRequest):
        if not data.pptx or not data.excel:
            return {"status": "error", "message": "Missing required parameters"}

        prs = Presentation(data.pptx)
        sample_num = len(prs.slides)

        data_by_sample = get_data_by_sample(data.excel)
        for i in data_by_sample.keys():
            if isinstance(i, str):
                self.log_warning(f"Sample number '{i}' is not an integer. Skipping sample '{i}'.")
                continue
            if i >= sample_num:
                self.log_warning(f"No sample slide with index {i} exists. Skipping sample {i}.\nHint: The 'Sample Num' column should start from 0 and be continuous.")
                continue
            try:
                SlideDrawer(prs, i, data_by_sample[i]).draw(
                    margin=(data.margin_x, data.margin_y),
                    padding=(data.padding_x, data.padding_y),
                    per_slide=data.per_slide
                )
            except Exception as e:
                return {"status": "error", "message": f"Error while drawing slide {i}: {str(e)}"},

        filename = filedialog.asksaveasfilename(
            defaultextension=".pptx",
            filetypes=[("PowerPoint files", "*.pptx")],
            title="Save the file as",
            initialfile=f"generated-{os.path.basename(data.pptx)}"
        )
        if filename:
            prs.save(filename)
            open_file_with_default_program(filename)
            return {"status": "success", "message": f"PPTX saved as '{os.path.basename(filename)}'"}
        else:
            return {"status": "success", "message": "Saving PPTX canceled by user"}

    def log_warning(self, message):
        if self.is_gui:
            messagebox.showwarning("Warning", message)
        logging.warning(message)


if __name__ == "__main__":
    args = get_args()
    task_manager = TaskManger(args.gui)

    if args.rpc:
        print("Python RPC mode ready")

        sys.stdout.reconfigure(line_buffering=True)
        for line in sys.stdin:
            try:
                request = json.loads(line.strip())
                task_manager.process_request(request)
            except Exception as e:
                print(json.dumps({"status": "error", "message": str(e)})) # task??
    else:
        data = GenerateRequest(
            pptx=args.pptx,
            excel=args.excel,
            margin_x=args.margin_x,
            margin_y=args.margin_y,
            padding_x=args.padding_x,
            padding_y=args.padding_y,
            per_slide=args.per_slide
        )
        result = task_manager.generate_pptx(data)
        print(result["message"])
