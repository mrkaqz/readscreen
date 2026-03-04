from ctypes import windll, Structure, c_long, byref
import time
import sys, os

class POINT(Structure):
    _fields_ = [("x", c_long), ("y", c_long)]


def queryMousePosition():
    pt = POINT()
    windll.user32.GetCursorPos(byref(pt))
    return { "x": pt.x, "y": pt.y}


def wait_key():
    ''' Wait for a key press on the console and return it. '''
    result = None
    if os.name == 'nt':
        import msvcrt
        result = msvcrt.getch()
    else:
        import termios
        fd = sys.stdin.fileno()

        oldterm = termios.tcgetattr(fd)
        newattr = termios.tcgetattr(fd)
        newattr[3] = newattr[3] & ~termios.ICANON & ~termios.ECHO
        termios.tcsetattr(fd, termios.TCSANOW, newattr)

        try:
            result = sys.stdin.read(1)
        except IOError:
            pass
        finally:
            termios.tcsetattr(fd, termios.TCSAFLUSH, oldterm)

    return result

scale_factor = 5
monitor = input('What is monitor number to record? (1/2/3).....')

print('Move mouse to point X1,Y1 and press any key')
wait_key()
pos = queryMousePosition()
x1 = pos['x']
y1 = pos['y']
print('X1: {}, Y1: {}'.format(x1,y1))

time.sleep(0.3) 

print('Move mouse to point X2,Y2 and press any key')
wait_key()
pos = queryMousePosition() 
x2 = pos['x']  
y2 = pos['y']
print('X2: {}, Y2: {}'.format(x2,y2))

time.sleep(0.3)     

print('------------------------------------\n\n')
print('{')
print('    "loc_x1" : "{}",'.format(x1))
print('    "loc_y1" : "{}",'.format(y1))
print('    "loc_x2" : "{}",'.format(x2))
print('    "loc_y2" : "{}",'.format(y2))
print('    "monitor" : "{}",'.format(monitor))
print('    "scale_factor" : "{}"'.format(scale_factor))
print('}  \n\n    ')


f = open("config.json", "w+")
f.write('{\n')
f.write('    "loc_x1" : "{}",\n'.format(x1))
f.write('    "loc_y1" : "{}",\n'.format(y1))
f.write('    "loc_x2" : "{}",\n'.format(x2))
f.write('    "loc_y2" : "{}",\n'.format(y2))
f.write('    "monitor" : "{}",\n'.format(monitor))
f.write('    "scale_factor" : "{}"\n'.format(scale_factor))
f.write('}')
f.close()

print('config.json file created')
wait_key()