from TableProcessing.parsing import get_conturs


def detect(path):

    myCells = get_conturs(path)

    myCells.save_to_file()
    myCells.save_to_excel()