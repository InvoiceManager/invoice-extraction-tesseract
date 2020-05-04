class Header:

    def __init__(self):
        self.list_of_proprieties = []
        self.ordered_list = []

    def add_proprieties(self, prop):
        self.list_of_proprieties.append(prop)

    def sort_list(self):
        self.list_of_proprieties.reverse()

    def print_props(self):
        print(self.list_of_proprieties)
