class Header:

    def __init__(self):
        self.list_of_proprieties = []
        self.ordered_list = []
        self.path_to_header_txt = "final/header_tabel.txt"

    def add_proprieties(self, prop):
        prop = prop.replace('\n', ' ').replace('\r', '')
        self.list_of_proprieties.append(prop)

    def sort_list(self):
        self.list_of_proprieties.reverse()

    def print_props(self):
        print(self.list_of_proprieties)

    def save_to_file(self):
        file = open(self.path_to_header_txt, "w", encoding="utf-8")
        for element in self.list_of_proprieties:
            file.write(element)
            file.write('\n')
        file.close()
