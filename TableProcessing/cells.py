import json
class Cells:

    def __init__(self):
        self.cells_dic = {}
        self.path_to_bodyr_txt = "final/body_tabel.txt"

    def add_cell(self, index, value):
        self.cells_dic[index] = value

    def print_dict(self):
        print(self.cells_dic)

    def save_to_file(self):
        file = open(self.path_to_bodyr_txt, "w")
        file.write(json.dumps(self.cells_dic))
        file.close()