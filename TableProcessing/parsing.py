import cv2

from TableProcessing import cells
from TableProcessing import header
from TextRecognition import tesseract


def sort_contours(cnts, method="left-to-right"):
    # initialize the reverse flag and sort index
    reverse = False
    i = 0
    # handle if we need to sort in reverse
    if method == "right-to-left" or method == "bottom-to-top":
        reverse = True
    # handle if we are sorting against the y-coordinate rather than
    # the x-coordinate of the bounding box
    if method == "top-to-bottom" or method == "bottom-to-top":
        i = 1
    # construct the list of bounding boxes and sort them from top to
    # bottom
    boundingBoxes = [cv2.boundingRect(c) for c in cnts]
    (cnts, boundingBoxes) = zip(*sorted(zip(cnts, boundingBoxes),
                                        key=lambda b: b[1][i], reverse=reverse))
    # return the list of sorted contours and bounding boxes
    return (cnts, boundingBoxes)


### MAKING TEMPLATE WITHOUT HOUGH
def get_conturs(file_name):
    # init header
    myHeader = header.Header()
    myCells = cells.Cells()
    # Read the image and make a copy then transform it to gray colorspace,
    # threshold the image and search for contours.
    img = cv2.imread(file_name)
    res = img.copy()
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    contours, hierarchy = cv2.findContours(thresh, cv2.RETR_TREE, cv2.CHAIN_APPROX_NONE)
    # Iterate through contours and draw a slightly bigger white rectangle
    # over the contours that are not big enough (the text) on the copy of the image.
    for i in contours:
        cnt = cv2.contourArea(i)
        if cnt < 500:
            x, y, w, h = cv2.boundingRect(i)
            cv2.rectangle(res, (x - 1, y - 1), (x + w + 1, y + h + 1), (255, 255, 255), -1)

    # Optional count the rows and columns of the table
    count = res.copy()
    gray = cv2.cvtColor(count, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    contours, hierarchy = cv2.findContours(thresh, cv2.RETR_TREE, cv2.CHAIN_APPROX_NONE)
    (contours, boundingBoxes) = sort_contours(contours, method="top-to-bottom")

    aux_line_x = 1
    line = []
    flag_to_header = False
    idx = 0
    aux_y = 0
    for i in contours:
        cnt = cv2.contourArea(i)
        if 50000 > cnt > 50:
            x, y, w, h = cv2.boundingRect(i)

            if idx == 0:
                aux_y = y
            idx += 1
            new_img = img[y:y + h, x:x + w]
            cv2.imwrite('TableProcessing/result/' + str(idx) + '.png', new_img)
            cv2.drawContours(count, [i], 0, (255, 255, 0), 2)

            extracted_text = tesseract.text_recognize('TableProcessing/result/' + str(idx) + '.png')
            extracted_text = extracted_text.replace('\n', ' ').replace('\r', '')
            if y - aux_y >= 10:
                flag_to_header = True

            if flag_to_header == False:
                myHeader.add_proprieties(extracted_text)
                aux_y = y
            else:
                if y - aux_y >= 10:
                    line.reverse()

                    myCells.add_cell(aux_line_x, line)
                    aux_line_x = aux_line_x + 1
                    line = []

                line.append(extracted_text)
                aux_y = y

            #  print(extracted_text)

    myHeader.sort_list()
    myHeader.save_to_file()
    return myCells
