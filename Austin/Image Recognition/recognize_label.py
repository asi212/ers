# import the necessary packages
import imutils
import numpy as np
import skimage
import pytesseract
import cv2
#import numpy as np
pytesseract.pytesseract.tesseract_cmd = r"C:\Users\a.ibele\AppData\Local\Tesseract-OCR\tesseract.exe"

dir = r"C:\Users\a.ibele\PycharmProjects\tensorflow"
path = r"C:\Users\a.ibele\PycharmProjects\tensorflow\chiller_1.jpg"
image = cv2.imread(path)
image = imutils.resize(image, height=500)

# load the query image, compute the ratio of the old height
# to the new height, clone it, and resize it
ratio = image.shape[0] / 300.0
orig = image.copy()
image = imutils.resize(image, height = 300)

# convert the image to grayscale, blur it, and find edges
# in the image
gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
gray = cv2.bilateralFilter(gray, 11, 17, 17)
edged = cv2.Canny(gray, 30, 200)

# find contours in the edged image, keep only the largest
# ones, and initialize our screen contour
cnts = cv2.findContours(edged.copy(), cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
cnts = imutils.grab_contours(cnts)
cnts = sorted(cnts, key = cv2.contourArea, reverse = True)[:10]
screenCnt = None

# loop over our contours
for c in cnts:
    # approximate the contour
    peri = cv2.arcLength(c, True)
    approx = cv2.approxPolyDP(c, 0.015 * peri, True)

    # if our approximated contour has four points, then
    # we can assume that we have found our label
    if len(approx) == 4:
        screenCnt = approx
        break


# now that we have our screen contour, we need to determine
# the top-left, top-right, bottom-right, and bottom-left
# points so that we can later warp the image -- we'll start
# by reshaping our contour to be our finals and initializing
# our output rectangle in top-left, top-right, bottom-right,
# and bottom-left order
pts = screenCnt.reshape(4, 2)
rect = np.zeros((4, 2), dtype="float32")

# the top-left point has the smallest sum whereas the
# bottom-right has the largest sum
s = pts.sum(axis=1)
rect[0] = pts[np.argmin(s)]
rect[2] = pts[np.argmax(s)]

# compute the difference between the points -- the top-right
# will have the minumum difference and the bottom-left will
# have the maximum difference
diff = np.diff(pts, axis=1)
rect[1] = pts[np.argmin(diff)]
rect[3] = pts[np.argmax(diff)]

# multiply the rectangle by the original ratio
rect *= ratio

# now that we have our rectangle of points, let's compute
# the width of our new image
(tl, tr, br, bl) = rect
widthA = np.sqrt(((br[0] - bl[0]) ** 2) + ((br[1] - bl[1]) ** 2))
widthB = np.sqrt(((tr[0] - tl[0]) ** 2) + ((tr[1] - tl[1]) ** 2))

# ...and now for the height of our new image
heightA = np.sqrt(((tr[0] - br[0]) ** 2) + ((tr[1] - br[1]) ** 2))
heightB = np.sqrt(((tl[0] - bl[0]) ** 2) + ((tl[1] - bl[1]) ** 2))

# take the maximum of the width and height values to reach
# our final dimensions
maxWidth = max(int(widthA), int(widthB))
maxHeight = max(int(heightA), int(heightB))

# construct our destination points which will be used to
# map the screen to a top-down, "birds eye" view
dst = np.array([
    [0, 0],
    [maxWidth - 1, 0],
    [maxWidth - 1, maxHeight - 1],
    [0, maxHeight - 1]], dtype="float32")

# calculate the perspective transform matrix and warp
# the perspective to grab the screen
M = cv2.getPerspectiveTransform(rect, dst)
warp = cv2.warpPerspective(orig, M, (maxWidth, maxHeight))

# convert the warped image to grayscale and then adjust
# the intensity of the pixels to have minimum and maximum
# values of 0 and 255, respectively
warp = cv2.cvtColor(warp, cv2.COLOR_BGR2GRAY)
warp = skimage.exposure.rescale_intensity(warp, out_range=(0, 255))
warp = imutils.resize(warp, height=500)


# crop desired regions

type = warp[145:178, 163:483]
partno = warp[200:233, 163:483]
snno = warp[255:288, 163:483]

cv2.imwrite( dir + r"\type.jpg", type )
cv2.imwrite( dir + r"\partno.jpg", partno )
cv2.imwrite( dir + r"\snno.jpg", snno )

#cv2.imshow("type", type)
#cv2.imshow("part no", partno)
#cv2.imshow("serial no", snno)
#cv2.waitKey(0)


# read text
# Method 1.. return = '219034734 S'
#image = cv2.imread(imgpath, cv2.IMREAD_GRAYSCALE)
image = snno
image = cv2.dilate(image, (5, 5), image)
text = pytesseract.image_to_string(image, config='--psm 7')


cv2.imshow('image', image)
cv2.waitKey(0)

print(text)

#image.save(dir + r"\rec_snno.jpg")

