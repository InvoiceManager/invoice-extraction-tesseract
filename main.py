from HeaderRecognition import recognitionTypeHeader
from TableProcessing import detectTableWithCv2
import cv2

# file = r'C:\Users\Andrada\OneDrive\Desktop\Master\Sem2\Statistica NLP\invoice-extraction-tesseract\src\tabel.png'

jsonFile = r'C:\Users\cosmin\PycharmProjects\9iuninvoice\final\body_tabel.json'

file = r'C:\Users\cosmin\PycharmProjects\9iuninvoice\src\tabel.png'
cr2 = r'C:\Users\cosmin\PycharmProjects\9iuninvoice\src\crop2.jpg'

header_final = r'C:\Users\cosmin\PycharmProjects\9iuninvoice\final\fact.txt'
if __name__ == '__main__':

    # extract header and type

    type = recognitionTypeHeader.getType(file)

    detectTableWithCv2.detect(cr2)

    header_data = recognitionTypeHeader.getContent(file, header_final, type)

    sheet_name = header_data[4].replace(" ", "_") + "-" + header_data[0].replace(" ", "_")
    recognitionTypeHeader.writeExcelHeader(header_data, 'Output_Invoices.xlsx', sheet_name)
    recognitionTypeHeader.writeExcelTable(jsonFile, 'Output_Invoices.xlsx', sheet_name)

    print("Nr Fact:", header_data[0])
    print("Seria Fact:", header_data[1])
    print("Data emiterii:", header_data[2])
    print("Data scad: ", header_data[3])
    print("Nume Furn:", header_data[4])
    print("Addresa furnizor: ", header_data[5])
    print("Cont furn: ", header_data[6])
    print("Banca Furn:", header_data[7])
    print("Nume Cump:", header_data[8])
    print("Adresa Cump: ", header_data[9])
    print("Cont Cump:", header_data[10])
    print("Banca Cump:", header_data[11])
    print("Total factura:", header_data[12])
