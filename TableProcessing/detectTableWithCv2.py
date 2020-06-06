from TableProcessing.parsing import get_conturs


def detect(type):
    file = r'C:\Faculty\Master1\Invoice\invoice-extraction-tesseract\src\capture3.png'

    if type == "cubus":
        myCells = get_conturs(file)

        myCells.save_to_file()
