import re
from argparse import Namespace
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkinterdnd2 import DND_FILES, TkinterDnD

def get_args_by_gui(args = Namespace(excel=None, pptx=None, margin_x=0.0, margin_y=0.0, padding_x=0.0, padding_y=0.0, per_slide=None)):
    root = TkinterDnD.Tk()
    root.title("Nametag Generator")
    root.geometry("600x350")  # Set initial window size
    
    def select_excel_file():
        excel_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        set_excel_file(excel_file)

    def set_excel_file(file_path):
        if file_path:
            args.excel = file_path
            excel_var.set("..." + file_path[-30:])
    
    def select_pptx_file():
        pptx_file = filedialog.askopenfilename(filetypes=[("PowerPoint files", "*.pptx")])
        set_pptx_file(pptx_file)

    def set_pptx_file(file_path):
        if file_path:
            args.pptx = file_path
            pptx_var.set("..." + file_path[-30:])
    
    def generate_nametags():
        if not args.excel or not args.pptx:
            text = "Please select both excel and pptx file"
            warning_label.config(text=text)
            print(text)
            return
        per_slide = per_slide_entry.get()
        if per_slide != 'max' and not per_slide.isdigit():
            text = "Invalid number of nametags per slide. Please enter a positive integer or 'max'"
            warning_label.config(text=text)
            print(text)
            return
        else:
            args.per_slide = int(per_slide) if per_slide != 'max' else None

        try:
            args.margin_x = float(margin_x_entry.get())
            args.margin_y = float(margin_y_entry.get())
            args.padding_x = float(padding_x_entry.get())
            args.padding_y = float(padding_y_entry.get())
        except ValueError:
            text = "Invalid margin or padding value. Please enter a number"
            warning_label.config(text=text)
            print(text)
            return

        print("Generating nametags...")
        print(f"Excel file: {args.excel}")
        print(f"PowerPoint file: {args.pptx}")
        print(f"Margin: ({args.margin_x}, {args.margin_y})")
        print(f"Padding: ({args.padding_x}, {args.padding_y})")
        print(f"Output per slide: {args.per_slide}")
        root.destroy()
    
    def on_enter(event):
        event.widget.configure(style='hover.TButton')
    
    def on_leave(event):
        event.widget.configure(style='default.TButton')

    def on_closing():
        import sys
        sys.exit(0)

    def on_drop(event):
        def to_list(data):
            matches = re.findall(r'\{(.*?)\}', data)
            for match in matches:
                data = data.replace("{" + match + "}", "")
            return list(matches) + data.split()

        files = to_list(event.data)
        for file in files:
            if file.endswith('.xlsx'):
                set_excel_file(file)
            elif file.endswith('.pptx'):
                set_pptx_file(file)
            else:
                print("Unsupported file format:", file)
        
        
    root.protocol("WM_DELETE_WINDOW", on_closing)

    root.drop_target_register(DND_FILES)
    root.dnd_bind('<<Drop>>', on_drop)

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

    # Padding inputs
    padding_frame = tk.Frame(root)
    padding_frame.pack(pady=10)

    padding_label = tk.Label(padding_frame, text="Padding")
    padding_label.grid(row=0, column=0, padx=8)

    padding_x_label = tk.Label(padding_frame, text="x:")
    padding_x_label.grid(row=0, column=1)
    padding_x_entry = tk.Entry(padding_frame, width=5)
    padding_x_entry.grid(row=0, column=2)
    padding_x_entry.insert(0, '0.0')
    padding_x_unit = tk.Label(padding_frame, text="cm")
    padding_x_unit.grid(row=0, column=3, padx=5)

    padding_y_label = tk.Label(padding_frame, text="y:")
    padding_y_label.grid(row=0, column=4)
    padding_y_entry = tk.Entry(padding_frame, width=5)
    padding_y_entry.grid(row=0, column=5)
    padding_y_entry.insert(0, '0.0')
    padding_y_unit = tk.Label(padding_frame, text="cm")
    padding_y_unit.grid(row=0, column=6, padx=5)

    # Margin inputs
    margin_frame = tk.Frame(root)
    margin_frame.pack(pady=10)

    margin_label = tk.Label(margin_frame, text="Margin")
    margin_label.grid(row=0, column=0, padx=8)

    margin_x_label = tk.Label(margin_frame, text="x:")
    margin_x_label.grid(row=0, column=1)
    margin_x_entry = tk.Entry(margin_frame, width=5)
    margin_x_entry.grid(row=0, column=2)
    margin_x_entry.insert(0, '0.0')
    margin_x_unit = tk.Label(margin_frame, text="cm")
    margin_x_unit.grid(row=0, column=3, padx=5)

    margin_y_label = tk.Label(margin_frame, text="y:")
    margin_y_label.grid(row=0, column=4)
    margin_y_entry = tk.Entry(margin_frame, width=5)
    margin_y_entry.grid(row=0, column=5)
    margin_y_entry.insert(0, '0.0')
    margin_y_unit = tk.Label(margin_frame, text="cm")
    margin_y_unit.grid(row=0, column=6, padx=5)
    
    # Output per slide input
    per_slide_frame = tk.Frame(root)
    per_slide_frame.pack(pady=10)

    per_slide_label = tk.Label(per_slide_frame, text="Output per slide:")
    per_slide_label.grid(row=0, column=0)
    per_slide_entry = tk.Entry(per_slide_frame, width=5)
    per_slide_entry.grid(row=0, column=1)
    per_slide_entry.insert(0, 'max')


    generate_button = ttk.Button(root, text="Generate Nametags", command=generate_nametags, style='default.TButton')
    generate_button.pack(pady=10)
    generate_button.bind("<Enter>", on_enter)
    generate_button.bind("<Leave>", on_leave)
    
    # Warning label for error messages
    warning_label = tk.Label(root, text="", fg="red")
    warning_label.pack(pady=10)

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
    print(get_args_by_gui())
