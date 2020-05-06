import pytesseract
import cv2
import numpy as np
import re
import openpyxl


def getContent(image_path, output_txt_path):
    ################PREPROCESARE##############
    # opening - erosion followed by dilation
    def opening(img):
        kernel = np.ones((5, 5), np.uint8)
        return cv2.morphologyEx(img, cv2.MORPH_OPEN, kernel)

    # canny edge detection
    def canny(img):
        return cv2.Canny(img, 100, 200)

    # get grayscale image
    def get_grayscale(img):
        return cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # noise removal
    def remove_noise(img):
        return cv2.medianBlur(img, 5)

    # thresholding
    def thresholding(img):
        return cv2.threshold(img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]

    # dilation
    def dilate(img):
        kernel = np.ones((5, 5), np.uint8)
        return cv2.dilate(img, kernel, iterations=1)

    # erosion
    def erode(img):
        kernel = np.ones((5, 5), np.uint8)
        return cv2.erode(img, kernel, iterations=1)

    # skew correction
    def deskew(img):
        coords = np.column_stack(np.where(img > 0))
        angle = cv2.minAreaRect(coords)[-1]
        if angle < -45:
            angle = -(90 + angle)
        else:
            angle = -angle
            (h, w) = img.shape[:2]
            center = (w // 2, h // 2)
            M = cv2.getRotationMatrix2D(center, angle, 1.0)
            rotated = cv2.warpAffine(img, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
        return rotated

    # template matching
    def match_template(img, template):
        return cv2.matchTemplate(img, template, cv2.TM_CCOEFF_NORMED)

    ######## HEADER ######
    def header_eon(header_txt):
        print("eon")

    def header_general(header_txt):
        antet = ''
        file1 = open(header_txt, "r")
        lines = file1.readlines()
        for line in lines:
            antet += line

        proper_noun = '(([A-Z][A-Za-z]{1,15}[ -.])([A-Z][A-Za-z]{1,15}[ -.]*)+)[\n]'
        address = '((Bulevardul|Bd.|Str.|Str|Strada)[: .,0-9a-zA-Z]+)'
        cont = '((RO)[0-9A-Z]{22})'
        banca = '^(Banca)[ 0-9A-Za-z]+'
        nr_factura = '((nr.|nr|Nr|Nr.|Numarul|Numărul|numarul|numărul)[ :][0-9]+)'
        data = '(([0-9]{2})[/.-]([0-9]{2})[/.-][0-9]{4})'
        seria = '((Seria|seria|Serie|serie)[: -]+[A-Z0-9a-z])'

        date_factura = ''
        numar_factura = ''
        seria_factura = ''
        data_emiterii = ''
        data_scadenta = ''
        nume_furnizor = ''
        date_furnizor = ''
        address_furnizor = ''
        cont_furnizor = ''
        banca_furnizor = ''
        nume_cumparator = ''
        date_cumparator = ''
        address_cumparator = ''
        cont_cumparator = ''
        banca_cumparator = ''
        denumire_produs = ''
        cantitate_produs = ''
        valoare_produs = ''
        total_factura = ''

        # luam datele facturii
        count_factura = 0
        for line in lines:
            count_factura += 1
            if re.search('factura|Factura|FACTURA|FACTURĂ|Factură', line):
                break

        with open(header_txt) as myantet:
            f_lines = myantet.readlines()[(count_factura - 2):(count_factura + 2)]
            for line in f_lines:
                date_factura += line

            numar_factura = re.findall(nr_factura, date_factura)[0][0]
            numar_factura = re.findall(r'([0-9]+)', numar_factura)[0]
            seria_factura = re.findall(seria, date_factura)[0][0]
            seria_factura = seria_factura.split(' ')[1]

            date = re.findall(data, date_factura)
            data_emiterii = date[0][0]
            if len(date) > 1:
                data_scadenta = date[1][0]
            # print(numar_factura)
            # print(data_emiterii)
            # print(data_scadenta)
        list_proper_noun = re.findall(proper_noun, antet)
        nume_furnizor = list_proper_noun[0][0]

        # unde incep datele furnizorului
        line_furnizor = 0
        for line in lines:
            line_furnizor = line_furnizor + 1
            if re.search(nume_furnizor, line):
                break

            # unde incep datele cumparatorului
            line_cumparator = 0

            # luam datele furnizorului
            with open("final/antet.txt") as myantet:
                f_lines = myantet.readlines()[line_furnizor:(line_furnizor + 10)]
                for line in f_lines:
                    line_cumparator = line_cumparator + 1
                    date_furnizor += line
                    if re.search(address, line):
                        address_furnizor = re.findall(address, line)[0][0]
                    if re.search(cont, line):
                        cont_furnizor = re.findall(cont, line)[0][0]
                    if re.search(banca, line):
                        banca_furnizor = line
                        break
                line_cumparator = line_cumparator + line_furnizor

            # luam datele cumparatorului
            with open("final/antet.txt") as myantet:
                f_lines = myantet.readlines()[line_cumparator:-1]
                for line in f_lines:
                    date_cumparator += line
                    if re.search(address, line):
                        address_cumparator = re.findall(address, line)[0][0]
                    if re.search(cont, line):
                        cont_cumparator = re.findall(cont, line)[0][0]
                    if re.search(banca, line):
                        banca_cumparator = line
                        break
                nume_cumparator = re.findall(proper_noun, date_cumparator)[0][0]
                # print(nume_cumparator)
                # print(address_cumparator)
                # print(cont_cumparator)
                # print(banca_cumparator)

        file1.close()
        print("Data fact:", date_factura)
        print("Nr Fact:", numar_factura)
        print("Seria Fact:", seria_factura)
        print("Data emiterii:", data_emiterii)
        print("Data scad: ", data_scadenta)
        print("Nume Furn:", nume_furnizor)
        print("Addresa furnizor: ", address_furnizor)
        print("Cont furn: ", cont_furnizor)
        print("Banca Furn:", banca_furnizor)
        print("Nume Cump:", nume_cumparator)
        print("Adresa Cump: ", address_cumparator)
        print("Cont Cump:", cont_cumparator)
        print("Banca Cump:", banca_cumparator)

    ####################
    image = cv2.imread(image_path)
    gray = get_grayscale(image)
    thresh = thresholding(gray)
    opening = opening(gray)
    canny = canny(gray)

    cv2.imshow("image", thresh)
    cv2.waitKey(0)

    # Adding custom options
    custom_config = r'-l ron --oem 3 --psm 3'
    text = pytesseract.image_to_string(image, config=custom_config)
    # print(text)

    text.replace('\n\n', '\n').replace('\n\n', '\n')
    file = open(output_txt_path, "w+", errors="ignore")
    file.write(text)
    # file.seek(0)
    # file.truncate()
    file.close()

    file = open(output_txt_path, "r")
    lines = file.readlines()
    count = 0
    for line in lines:
        file_antet = open("final/antet.txt", "a")
        if count > 5:
            break
        if len(line.strip()) == 0:
            count = count + 1
        else:
            count = 0
            file_antet.write(line)
    file_antet.close()
    with open("final/antet.txt") as f:
        if 'E.ON' in f.read() or 'e-on' in f.read() or 'energie electric' in f.read():
            header_eon("antet.txt")
            type = "eon"
        else:
            header_general("final/antet.txt")
            type = "others"
    return type