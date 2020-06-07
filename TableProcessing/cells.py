import json
import xlwt
from xlwt import Workbook



class Cells:

    def __init__(self):
        self.cells_dic = {}
        self.path_to_bodyr_txt = "final/body_tabel.json"
        self.header = []
        # Workbook is created
        self.wb = Workbook()

        # add_sheet is used to create sheet.
        self.sheet1 = self.wb.add_sheet('Sheet 1')

    def add_header(self, list):
        self.header = list

    def add_cell(self, index, value):
        self.cells_dic[index] = value

    def print_dict(self):
        print(self.cells_dic)

    def save_to_excel(self):
        for key in self.cells_dic.keys():
            list = self.cells_dic.get(key)
            for i in range(0, len(list)):
                self.sheet1.write(key, i, list[i])

        self.wb.save('xlwt example.xls')
    def save_to_file(self):
        self.cells_dic[1] = self.header
        with open(self.path_to_bodyr_txt, 'w', encoding="utf-8") as fp:
            json.dump(self.cells_dic, fp)
