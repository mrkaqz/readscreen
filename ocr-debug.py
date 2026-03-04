import pytesseract as tess
tess.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
from PIL import Image, ImageFilter, ImageGrab
import cv2
import os

#os.system('color f0') # sets the background to white
tess_config = '--psm 6 --oem 1'
#tess_config = ''

scale_factor = input('What is image resize scale (1-10)? ')
if int(scale_factor) in range(1,10):
    scale_factor = int(scale_factor)
else:
    print("Use default scale = 1")
    scale_factor = 1

img = cv2.imread('screenshot.png')
img = cv2.resize(img, None, fx=scale_factor, fy=scale_factor, interpolation=cv2.INTER_CUBIC)
cv2.imwrite('screenshot_resize.png', img)

while True:


    img = cv2.imread('screenshot_resize.png')
    text = tess.image_to_string(img, config=tess_config)
    #text = tess.image_to_string(img, config='--psm 10 --oem 1')
    print('---------- original ---------')
    print(text)


    img = Image.open("screenshot_resize.png")
    newimgdata = []
    white = (255,255,255)
    black = (1,1,1)

    for color in img.getdata():
        if color[0] in range(100,150):
            newimgdata.append( white )
        else:
            newimgdata.append( black )
    newimg = Image.new(img.mode,img.size)
    newimg.putdata(newimgdata)
    #newimg.show()
    newimg.save('screenshot_replace.png')

    text = tess.image_to_string(newimg, config=tess_config)
    print('---------- replace ---------')
    print(text)

    img = cv2.imread('screenshot_resize.png')
    gry = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    thr = cv2.threshold(gry, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    text = tess.image_to_string(thr, config=tess_config)
    cv2.imwrite('screenshot_resize_thr.png', thr)
    print('---------- threshold ---------')
    print(text)

    input("Press Enter to continue...\n")

    
