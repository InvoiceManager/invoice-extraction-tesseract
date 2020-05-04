class Cells:

    def __init__(self):
        self.cells_dic = {}

    def add_cell(self, index, value):
        self.cells_dic[index] = value

    def print_dict(self):
        print(self.cells_dic)