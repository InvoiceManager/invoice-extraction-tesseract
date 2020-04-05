import cv2
import numpy as np
import matplotlib.pyplot as plt

file = r'src/test.png'

im1 = cv2.imread(file, 0)
im = cv2.imread(file)
height, width, channels = im.shape
ret, thresh_value = cv2.threshold(im1, 180, 255, cv2.THRESH_BINARY_INV)

kernel = np.ones((5, 5), np.uint8)
dilated_value = cv2.dilate(thresh_value, kernel, iterations=1)

contours, hierarchy = cv2.findContours(dilated_value, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
cordinates = []
for cnt in contours:
    x, y, w, h = cv2.boundingRect(cnt)
    cordinates.append((x, y, w, h))
    # bounding the images
    if y < height  and h-x >150 and w-y >150 and h-x < height-50:
        cv2.rectangle(im, (x, y), (x + w, y + h), (0, 0, 255), 1)

plt.imshow(im)
cv2.namedWindow('detecttable2', cv2.WINDOW_NORMAL)
cv2.imwrite('img/detecttable2.jpg', im)