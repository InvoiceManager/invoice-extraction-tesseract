import cv2
import numpy as np

from TableProcessing import header
from TextRecognition import tesseract
from TableProcessing import cells

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


def extractCells(file_name):
    # init header
    myHeader = header.Header()
    myCells = cells.Cells()

    # Read the image
    img = cv2.imread(file_name, 0)

    # Thresholding the image
    (thresh, img_bin) = cv2.threshold(img, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)  # Invert the image
    img_bin = 255 - img_bin

    # Defining a kernel length
    kernel_length = np.array(img).shape[1] // 80

    # A verticle kernel of (1 X kernel_length), which will detect all the verticle lines from the image.
    verticle_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1,
                                                                 kernel_length))  # A horizontal kernel of (kernel_length X 1), which will help to detect all the horizontal line from the image.
    hori_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (kernel_length, 1))  # A kernel of (3 X 3) ones.
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))

    # Morphological operation to detect vertical lines from an image
    img_temp1 = cv2.erode(img_bin, verticle_kernel, iterations=3)
    verticle_lines_img = cv2.dilate(img_temp1, verticle_kernel, iterations=3)

    img_temp2 = cv2.erode(img_bin, hori_kernel, iterations=3)
    horizontal_lines_img = cv2.dilate(img_temp2, hori_kernel, iterations=3)

    # Weighting parameters, this will decide the quantity of an image to be added to make a new image.
    alpha = 0.5
    beta = 1.0 - alpha
    img_final_bin = cv2.addWeighted(verticle_lines_img, alpha, horizontal_lines_img, beta, 0.0)
    img_final_bin = cv2.erode(~img_final_bin, kernel, iterations=2)
    (thresh, img_final_bin) = cv2.threshold(img_final_bin, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
    cv2.imwrite("TableProcessing/result/img_final_bin.jpg", img_final_bin)

    # Find contours for image, which will detect all the boxes
    contours, hierarchy = cv2.findContours(img_final_bin, cv2.RETR_TREE,
                                           cv2.CHAIN_APPROX_SIMPLE)  # Sort all the contours by top to bottom.
    (contours, boundingBoxes) = sort_contours(contours, method="top-to-bottom")

    aux_header_x = -1
    aux_line_x = -1
    lines = []
    flag_to_header = False
    idx = 0
    for c in contours:
        # Returns the location and width,height for every contour
        x, y, w, h = cv2.boundingRect(c)
        if aux_header_x == x:
            aux_line_x = x
            flag_to_header = True
            if flag_to_header:
                aux_header_x = -2
                myHeader.sort_list()
                #TODO return header with body in one call
                myHeader.print_props()
                myHeader.save_to_file()

        if aux_header_x == -1:
            aux_header_x = x

        idx += 1
        new_img = img[y - 2:y + h + 2, x - 2:x + w]
        # If the box height is greater then 20, widht is >80, then only save it as a box in "cropped/" folder.
        cv2.imwrite('TableProcessing/result/' + str(idx) + '.png', new_img)

        extracted_text = tesseract.text_recognize('TableProcessing/result/' + str(idx) + '.png')
        if flag_to_header == False:
            myHeader.add_proprieties(extracted_text)
        if flag_to_header:
            if aux_line_x == x:
                myCells.add_cell(idx, lines)
                lines = []
            lines.append(extracted_text)

    return myCells