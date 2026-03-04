from PIL import ImageGrab, Image
import win32gui

while True:
    name  = win32gui.GetWindowText (win32gui.GetForegroundWindow())
    print(name)