import subprocess, os, platform

from collections import defaultdict
from openpyxl import load_workbook

class dotdict(dict):
    """dot.notation access to dictionary attributes"""
    __getattr__ = dict.get
    __setattr__ = dict.__setitem__
    __delattr__ = dict.__delitem__
    def copy(self):
        return dotdict(super().copy())

def chunk_list(l, chunk_size):
    return [l[i:i + chunk_size] for i in range(0, len(l), chunk_size)]

def tuples_to_dict_list(header, data):
    return [dict(zip(header, d)) for d in data]

def read_excel_data(filename):
    workbook = load_workbook(filename, data_only=True)
    sheet = workbook.active
    data = []
    for row in sheet.iter_rows(values_only=True):
        data.append(row)

    assert len(data) > 0, "No data found in the excel file"
    assert len(data) > 1, "Only header found in the excel file. No data found"
    for i, row in enumerate(data):
        data[i] = tuple(str(c) if c is not None else "" for c in row)

    header = data[0]
    header = [h.strip().lower() for h in header]
        
    return header, data[1:]

def headed_data_with_sample_num(header, data):
    try:
        sample_num_idx = header.index("sample num")
    except:
        sample_num_idx = len(header)
        header += ("sample num",)
        data = [d + (0,) for d in data]
    else:
        for i, d in enumerate(data):
            d = list(d)
            if isinstance(d[sample_num_idx], str):
                if d[sample_num_idx].strip() == "":
                    continue
                try:
                    d[sample_num_idx] = int(d[sample_num_idx].strip())
                except:
                    raise ValueError(f"Sample number '{d[sample_num_idx]}' is not an integer in row {i+2}")
            data[i] = tuple(d)
    data = tuples_to_dict_list(header, data)
    return data

def group_by_sample(data):
    data_by_sample = defaultdict(list)
    for d in data:
        data_by_sample[d["sample num"]].append(d)
    return data_by_sample

def get_data_by_sample(filename):
    header, data = read_excel_data(filename)
    haeded_data = headed_data_with_sample_num(header, data)
    return group_by_sample(haeded_data)

def open_file_with_default_program(filename):
    if os.path.isfile(filename):
        if platform.system() == 'Windows':
            os.startfile(filename)
        elif platform.system() == 'Darwin':  # macOS
            subprocess.run(('open', filename))
        else:  # linux variants
            subprocess.run(('xdg-open', filename))
    else:
        raise FileNotFoundError(f"No file found at {filename}")
    
def set_color(source_shape, target_shape):
    try:
        color = source_shape.fill.fore_color.rgb
        target_shape.fill.solid()
        target_shape.fill.fore_color.rgb = color
    except TypeError:
        pass
    except AttributeError:
        try:
            color = source_shape.fill.fore_color.theme_color
            target_shape.fill.solid()
            target_shape.fill.fore_color.theme_color = color
            target_shape.fill.fore_color.brightness = source_shape.fill.fore_color.brightness
        except TypeError:
            pass

def set_line(source_line, target_line): # _BasePicture, Shape, Connector
    target_line.dash_style = source_line.dash_style
    target_line.width = source_line.width
    set_color(source_line, target_line)

def set_shadow(source_shadow, target_shadow): # BaseShape
    if not source_shadow.inherit:
        ...
        # target_shadow.blur_radius = source_shadow.blur_radius
        # target_shadow.distance = source_shadow.distance
        # target_shadow.direction = source_shadow.direction
        # target_shadow.rotation = source_shadow.rotation
        # set_color(source_shadow.color, target_shadow.color)
        # target_shadow.visible = source_shadow.visible

def set_base_shape(source_shape, target_shape): # BaseShape
    target_shape.rotation = source_shape.rotation
    set_shadow(source_shape.shadow, target_shape.shadow)