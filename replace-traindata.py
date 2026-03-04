
import os
from elevate import elevate
from rich import print
import time

elevate()

print('Check Tesseract train data file >>> ',end = '')
time.sleep(1)
try:
    tessdata_size = os.path.getsize("C:\\Program Files\\Tesseract-OCR\\tessdata\eng.traineddata")
except:
    print('[red]train data not found[/red]')
else:
    if tessdata_size < 6000000:
        print('[yellow]basic[/yellow]')
    elif tessdata_size > 15000000:
        print('[green]best[/green]')
    elif tessdata_size > 22000000:
        print('[green]LSTM + Legacy[/green]')  

if tessdata_size < 6000000:
    copy_best = input('Training data is basic, want to use best data set instead? (y/n)')
    if copy_best.lower() == 'y':
            os.system('copy /Y "eng.traineddata" "C:\\Program Files\\Tesseract-OCR\\tessdata"')
            print('[green]copy completed[/green]')
    else:
        pass

exit = input('Enter to exit')