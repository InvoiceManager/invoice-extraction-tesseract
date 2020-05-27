from HeaderRecognition import recognitionTypeHeader
from TableProcessing import detectTableWithCv2
import cv2

file = r'C:\Users\cosmin\PycharmProjects\27maiinvoice\src\eon2.jpg'
cr1 = r'C:\Users\cosmin\PycharmProjects\27maiinvoice\src\crop1.jpg'
cr2 = r'C:\Users\cosmin\PycharmProjects\27maiinvoice\src\crop2.jpg'

header_final = r'C:\Users\cosmin\PycharmProjects\27maiinvoice\final\fact.txt'

if __name__ == '__main__':
    # extract header

    type = recognitionTypeHeader.getType(file)
    print(type)
    image1 = cv2.imread(file)
    if type == "eon":
        crop1 = image1[0:1065, 0:3000]
        cv2.imwrite(cr1, crop1)
        crop2 = image1[1060:2010, 0:3000]
        cv2.imwrite(cr2, crop2)
    elif type == "cubus":
        crop1 = image1[0:375, 0:2000]
        cv2.imwrite(cr1, crop1)
        crop2 = image1[370:1290, 0:2000]
        cv2.imwrite(cr2, crop2)
        detectTableWithCv2.detect(type)
    else:
        crop1 = image1[0:250, 0:2000]
        cv2.imwrite(cr1, crop1)
        crop2 = image1[245:2000, 0:2000]
        cv2.imwrite(cr2, crop2)

    header_data = recognitionTypeHeader.getContent(cr1, header_final, type)
    #recognitionTypeHeader.writeExcel(header_data, 'Output-Facturi.xlsx')

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

