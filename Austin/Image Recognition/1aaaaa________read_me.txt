This folder contains a few scripts for recognizing the label of a chiller from an image and extracting the serial number

2times_contour recognition_BW_.....py   is the best I could do thus far.
	- reads text image to text string using "tesseract" by google
	- takes about 30 seconds to extract text from 1 image TOOO SLOW

future work:
	- unlikely that above script can be optimized much
	- Possible solutions:
		- find a fastest OCR package than tesseract
		- Using machine learning to train an algorithm on the actual text from labels