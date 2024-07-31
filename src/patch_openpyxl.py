from openpyxl.cell import Cell
from openpyxl.worksheet._reader import WorksheetReader

def bind_cells(self):
    for idx, row in self.parser.parse():
        for cell in row:
            print(cell)
            print(self.ws.parent._cell_styles)
            try:
                style = self.ws.parent._cell_styles[cell['style_id']]
            except IndexError:
                style = None
            c = Cell(self.ws, row=cell['row'], column=cell['column'], style_array=style)
            c._value = cell['value']
            c.data_type = cell['data_type']
            self.ws._cells[(cell['row'], cell['column'])] = c
    self.ws.formula_attributes = self.parser.array_formulae
    if self.ws._cells:
        self.ws._current_row = self.ws.max_row # use cells not row dimensions
WorksheetReader.bind_cells = bind_cells
# See here for better solution: https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1673