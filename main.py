from HeaderRecognition import recognitionTypeHeader
from TableProcessing import detectTableWithCv2

file = r'C:\Faculty\Master1\Invoice\invoice-extraction-tesseract\src\fact.jpg'
header_final = r'C:\Faculty\Master1\Invoice\invoice-extraction-tesseract\final\fact.txt'

if __name__ == '__main__':
    # extract header
    recognitionTypeHeader.getContent(file, header_final)

    # tabel detect and recognize
    detectTableWithCv2.detect(file)
