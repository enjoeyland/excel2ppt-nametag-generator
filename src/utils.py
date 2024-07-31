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
        print(f"No file found at {filename}")