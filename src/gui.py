from argparse import Namespace
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk


def get_args_by_gui(args = Namespace(excel=None, pptx=None)):
    root = tk.Tk()
    root.title("Nametag Generator")
    root.geometry("600x200")  # Set initial window size
    
    def select_excel_file():
        excel_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if excel_file:
            args.excel = excel_file
            excel_var.set("..." + excel_file[-30:])
    
    def select_pptx_file():
        pptx_file = filedialog.askopenfilename(filetypes=[("PowerPoint files", "*.pptx")])
        if pptx_file:
            args.pptx = pptx_file
            pptx_var.set("..." + pptx_file[-30:])
    
    def generate_nametags():
        if args.excel and args.pptx:
            print(f"Excel file: {args.excel}")
            print(f"PowerPoint file: {args.pptx}")
            root.destroy()
        else:
            print("Please select both excel and pptx file")
    
    def on_enter(event):
        event.widget.configure(style='hover.TButton')
    
    def on_leave(event):
        event.widget.configure(style='default.TButton')

    frame1 = tk.Frame(root)
    frame1.pack(pady=10)
    
    excel_var = tk.StringVar(value="Select Excel File")
    excel_label = tk.Label(frame1, textvariable=excel_var)
    excel_label.pack(side=tk.LEFT)

    excel_button = ttk.Button(frame1, text="Select Excel File", command=select_excel_file, style='default.TButton')
    excel_button.pack(side=tk.LEFT, padx=10)
    excel_button.bind("<Enter>", on_enter)
    excel_button.bind("<Leave>", on_leave)
    
    excel_label = tk.Label(frame1, text="")
    excel_label.pack(side=tk.LEFT)
    
    frame2 = tk.Frame(root)
    frame2.pack(pady=10)
    
    pptx_var = tk.StringVar(value="Select PowerPoint File")
    pptx_label = tk.Label(frame2, textvariable=pptx_var)
    pptx_label.pack(side=tk.LEFT)

    pptx_button = ttk.Button(frame2, text="Select PowerPoint File", command=select_pptx_file, style='default.TButton')
    pptx_button.pack(side=tk.LEFT, padx=10)
    pptx_button.bind("<Enter>", on_enter)
    pptx_button.bind("<Leave>", on_leave)

    pptx_label = tk.Label(frame2, text="")
    pptx_label.pack(side=tk.LEFT)
    
    generate_button = ttk.Button(root, text="Generate Nametags", command=generate_nametags, style='default.TButton')
    generate_button.pack(pady=10)
    generate_button.bind("<Enter>", on_enter)
    generate_button.bind("<Leave>", on_leave)
    
    if args and args.excel:
        excel_var.set("..." + args.excel[-30:])
    elif args and args.pptx:
        pptx_var.set("..." + args.pptx[-30:])

    # Define a custom style for rounded buttons
    s = ttk.Style()
    s.configure('default.TButton', font=('Helvetica', 10), padding=5, borderwidth=2, relief="groove")
    s.map('default.TButton',
          foreground=[('active', 'black')],
          background=[('active', 'lightgray')])
    s.configure('hover.TButton', background='gray', font=('Helvetica', 10), padding=5, borderwidth=2, relief="groove")

    root.mainloop()

    return args

if __name__ == "__main__":
    get_args_by_gui()
