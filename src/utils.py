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