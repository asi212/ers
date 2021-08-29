# import the necessary packages
import imutils
import numpy as np
import skimage
import cv2
from imutils.perspective import four_point_transform

dir = r"C:\Users\a.ibele\PycharmProjects\tensorflow"
path = r"C:\Users\a.ibele\PycharmProjects\tensorflow\chiller_2.jpg"
image = cv2.imread(path)
#image = imutils.resize(image, height=500)

# load the query image, compute the ratio of the old height
# to the new height, clone it, and resize it
ratio = image.shape[0] / 600.0
orig = image.copy()
image = imutils.resize(image, height = 600)

# convert the image to grayscale, blur it, and find edges
# in the image
gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
gray = cv2.bilateralFilter(gray, 11, 17, 17)
edged = cv2.Canny(gray, 30, 200)

# find contours in the edged image, keep only the largest
# ones, and initialize our screen contour
cnts = cv2.findContours(edged.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE) # cv2.RETR_EXTERNAL because we only need external contours
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
snno = warp[245:305, 150:630]

# create edged image again
edged2 = cv2.Canny(snno, 50, 200, 255)

cnts2, hierarchy = cv2.findContours(edged2, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE) # cv2.RETR_EXTERNAL because we only need external contours
# Find the index of the largest contour
areas = [cv2.contourArea(c) for c in cnts2]
max_index = np.argmax(areas)
biggest_cnt = cnts2[max_index]


x, y, w, h = cv2.boundingRect(biggest_cnt)
y = y + 8
h = h - 14
x = x + 10
w = w - 190
cropped_snno = snno[y: y+h, x: x+w]


cropped_snno = cv2.threshold(cropped_snno, 127, 255, cv2.THRESH_BINARY)[1]
#cropped_snno2 = cv2.threshold(cropped_snno0,0,255,cv2.THRESH_BINARY+cv2.THRESH_OTSU)

image = cropped_snno
edged3 = cv2.Canny(cropped_snno, 50, 200, 255)
image_copy = image

contours, hier = cv2.findContours(image, cv2.RETR_LIST,cv2.CHAIN_APPROX_NONE)
for c in contours:
    area = cv2.contourArea(c)
    rect = cv2.minAreaRect(c)
    box = cv2.boxPoints(rect)
    # convert all coordinates floating point values to int
    box = np.int0(box)
    # draw a green 'nghien' rectangle
    if area > 1:
        cv2.drawContours(image_copy, [box], 0, (0, 255, 0), 1)
        #print([box])
cv2.imshow('image', image_copy)
cv2.waitKey(0)



image = cv2.drawContours(image, contours, -1, (255, 255, 0), 1)


cv2.imshow('image', cropped_snno)
cv2.waitKey(0)






### uses connected components with stats to select the characters -- did not work
#################
ret, thresh = cv2.threshold(cropped_snno,0,255,cv2.THRESH_BINARY+cv2.THRESH_OTSU)
# You need to choose 4 or 8 for connectivity type
connectivity = 4
# Perform the operation
output = cv2.connectedComponentsWithStats(thresh, connectivity, cv2.CV_32S)
# Get the results
# The first cell is the number of labels
num_labels = output[0]
# The second cell is the label matrix
labels = output[1]
# The third cell is the stat matrix. Contents = cv2.CC_STAT_LEFT,  cv2.CC_STAT_TOP, cv2.CC_STAT_WIDTH, cv2.CC_STAT_HEIGHT, cv2.CC_STAT_AREA
stats = output[2]
# The fourth cell is the centroid matrix
centroids = output[3]

i = 0
while i < len(stats):
    x = stats[i][0]
    y = stats[i][1]
    w = stats[i][2]
    h = stats[i][3]

pt1 = (stats[i][0], stats[i][1])
pt2 = (stats[i][0]+stats[i][2], stats[i][1]+stats[i][3])

box = cv2.rectangle(cropped_snno, pt1, pt2, color=(0, 255, 0), thickness=3)

cv2.imshow('box', box)
cv2.waitKey(0)

#stats[label, COLUMN]
#cv2.CC_STAT_LEFT, cv2.CC_STAT_TOP, cv2.CC_STAT_WIDTH, cv2.CC_STAT_HEIGHT)




