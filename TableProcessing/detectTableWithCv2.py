import cv2
import numpy as np
import matplotlib.pyplot as plt
from TableProcessing.cellExtractionCub import extractCells


def detect(type):
    cr2 = r'C:\Faculty\Master1\Invoice\invoice-extraction-tesseract\src\crop2.jpg'

    if type == "cubus":
        myCells = extractCells(cr2)
        myCells.save_to_file()

