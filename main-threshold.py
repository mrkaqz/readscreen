import mss
import mss.tools
import pytesseract as tess
tess.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
from PIL import Image, ImageFilter
import json
import csv
import time
import re
import cv2

def data_check(data_list):

    # add . if not exist
    for c in range(len(data_list)):
        d = data_list[c]
        if "." not in d:
            data_list[c] = f'{d[:len(d)-2]}.{d[len(d)-2:]}'
    
    #check data if range, if not return zero
    try:
        if float(data_list[1]) >= 100 or float(data_list[2]) >= 360:
            data_list = ['9.99','9.99','9.99']
    except:
        pass
    
    return data_list

#declear version
print('Read Screen Utility for Real Time Survey Calculation')
print('By Ronnarong Wongmalasit (ron@slb.com)')
print('Version: 0.3.1 Date: 1-May-20\n')

#read config file
print('Reading config file.....', end = '')
with open('config.json') as file:
    config_json = json.load(file)
    print('Done')

print('Press Ctrl-C to quit.\n')
scale_index = 0
try:

    while True:

        with open('config.json') as file:
            config_json = json.load(file)

        loc_x1 = int(config_json['loc_x1'])
        loc_x2 = int(config_json['loc_x2'])
        loc_y1 = int(config_json['loc_y1'])
        loc_y2 = int(config_json['loc_y2'])
        monitor_number = int(config_json['monitor'])
        scale_factor = float(config_json['scale_factor'])

        #take screenshot of realtime table
        with mss.mss() as sct:
            # Get information of monitor X
            mon = sct.monitors[monitor_number]
            # The screen part to capture
            monitor = {
                "top":loc_y1,  # px from the top
                "left": loc_x1,  # px from the left
                "width": loc_x2-loc_x1,
                "height": loc_y2-loc_y1,
                "mon": monitor_number,
            }
            output = "screenshot.png".format(**monitor)
            # Grab the data
            sct_img = sct.grab(monitor)
            # Save to the picture file
            mss.tools.to_png(sct_img.rgb, sct_img.size, output=output)



        #read text out of image
        for i in range(10):
            try:
                scale_factor = 1
                img = cv2.imread('screenshot.png')
                img = cv2.resize(img, None, fx=scale_factor, fy=scale_factor, interpolation=cv2.INTER_CUBIC)
                gry = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
                thr = cv2.threshold(gry, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
                text = tess.image_to_string(thr)
                cv2.imwrite('screenshot_resize_thr.png', thr)
                #print(text)

            except:
                print('Read Image Error -- Retrying...')
                continue
            else:
                break            
        
        #clean string 
        text = re.sub(r'[^\d\.MR\n]', '', text)
        text = text.replace('.M','M')
        text = text.replace('.R','R')
        text = text.replace('R.','M')
        text = text.replace('R.','R')
        list_text = text.split('\n')

        mwd_list = []
        rss_list = []
        i=0
        for item in list_text:
            if i==0:
                mwd_list = item.split('M')
            elif i==1:
                rss_list = item.split('R')
            i+=1

        if len(mwd_list) != 3:
            mwd_list = ['0.00','0.00','0.00']
        if len(rss_list) != 3:
            rss_list = ['0.00','0.00','0.00']



        #perform data check
        mwd_list = data_check(mwd_list)
        rss_list = data_check(rss_list)


        if mwd_list[0] == '9.99':
            mwd_out = ['Reading Out Of Range']
            mwd_list = ['0.00','0.00','0.00']
        elif mwd_list[0] == '':
            mwd_out = ['No Data Found']
            mwd_list = ['0.00','0.00','0.00']
        elif mwd_list[0] == '0.00':
            mwd_out = ['Reading Error']
        else:
            try:
                float(mwd_list[0])
                float(mwd_list[1])
                float(mwd_list[2])
                mwd_out = mwd_list
            except:
                mwd_out = ['Not a Number']
                mwd_list = ['0.00','0.00','0.00']
            

        if rss_list[0] == '9.99':
            rss_out = ['Reading Out Of Range']
            rss_list = ['0.00','0.00','0.00']
        elif rss_list[0] == '':
            rss_out = ['No Data Found']
            rss_list = ['0.00','0.00','0.00']
        elif rss_list[0] == '0.00':
            rss_out = ['Reading Error']
        else:
            try:
                float(rss_list[0])
                float(rss_list[1])
                float(rss_list[2])
                rss_out = rss_list
            except:
                rss_out = ['Not a Number']
                rss_list = ['0.00','0.00','0.00']


        print(mwd_out)  
        print(rss_out)
        print('---------------------------')

        #output  csv file
        for i in range(10):
            try:
                with open('output.csv', mode='w', newline='') as output_file:
                    output_writer = csv.writer(output_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                    output_writer.writerow(mwd_list)
                    output_writer.writerow(rss_list)
            except:
                print('Output to file error -- Retrying...')
                continue
            else:
                break

        time.sleep(3)

except KeyboardInterrupt:
    pass