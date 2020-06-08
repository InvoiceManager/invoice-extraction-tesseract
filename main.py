from HeaderRecognition import recognitionTypeHeader
from TableProcessing import detectTableWithCv2
import cv2

file = r'C:\Faculty\Master1\Invoice\invoice-extraction-tesseract\src\tabel.png'
cr2 = r'C:\Faculty\Master1\Invoice\invoice-extraction-tesseract\src\crop2.jpg'
jsonFile = r'C:\Faculty\Master1\Invoice\invoice-extraction-tesseract\final\body_tabel.json'


header_final = r'C:\Faculty\Master1\Invoice\invoice-extraction-tesseract\final\fact.txt'
if __name__ == '__main__':
    # extract header

    type = recognitionTypeHeader.getType(file)
    print(type)
    image1 = cv2.imread(file)
    if type == "eon":
        crop2 = image1[1060:2010, 0:3000]
        cv2.imwrite(cr2, crop2)
    elif type == "cubus":
        crop2 = image1[370:1290, 0:2000]
        cv2.imwrite(cr2, crop2)
        detectTableWithCv2.detect(type)
    else:
        crop2 = image1[245:2000, 0:2000]
        cv2.imwrite(cr2, crop2)

    header_data = recognitionTypeHeader.getContent(file, header_final, type)
    
    #Aici trebuie folosite datele intoarse de Cosmin din interfata (in loc de header_data)
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
