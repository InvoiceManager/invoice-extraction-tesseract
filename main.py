from HeaderRecognition import recognitionTypeHeader
from TableProcessing import detectTableWithCv2

file = r'C:\Faculty\Master1\Invoice\invoice-extraction-tesseract\src\fact.jpg'
header_final = r'C:\Faculty\Master1\Invoice\invoice-extraction-tesseract\final\fact.txt'

if __name__ == '__main__':
    # extract header
    
    type, header_data = recognitionTypeHeader.getContent(file, header_final)
    recognitionTypeHeader.writeExcel(header_data,'Output-Facturi.xlsx')
    
    print(type)
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
    print("Cont Cump:", header_data[11])
    print("Banca Cump:", header_data[11])



    # tabel detect and recognize
    detectTableWithCv2.detect(file)

