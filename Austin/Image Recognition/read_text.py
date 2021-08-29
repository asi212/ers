import pytesseract
#from PIL import Image, ImageEnhance, ImageFilter
import cv2
#import numpy as np
pytesseract.pytesseract.tesseract_cmd = r"C:\Users\a.ibele\AppData\Local\Tesseract-OCR\tesseract.exe"


dir = r"C:\Users\a.ibele\PycharmProjects\tensorflow"
imgpath = r"C:\Users\a.ibele\PycharmProjects\tensorflow\snno.jpg"


# Method 1.. return = '219034734 S'
image = cv2.imread(imgpath, cv2.IMREAD_GRAYSCALE)
image2 = cv2.dilate(image, (5, 5), image)
text = pytesseract.image_to_string(image, config='--psm 7')


cv2.imshow('image', image)
cv2.imshow('image2', image2)
cv2.waitKey(0)


image.save(dir + r"\rec_snno.jpg")

