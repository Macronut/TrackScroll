import sys
import usb.core
import usb.util
import win32api
import win32con
import time
import math
from ctypes import *

LEFTHAND = True
ACCELERATE = 0.25
ROTATION = 0.15

MOUSEEVENTF_WHEEL = win32con.MOUSEEVENTF_WHEEL
MOUSEEVENTF_HWHEEL = 0x1000
MOUSEEVENTF_LEFTDOWN = win32con.MOUSEEVENTF_LEFTDOWN
MOUSEEVENTF_LEFTUP = win32con.MOUSEEVENTF_LEFTUP
MOUSEEVENTF_RIGHTDOWN = win32con.MOUSEEVENTF_RIGHTDOWN
MOUSEEVENTF_RIGHTUP = win32con.MOUSEEVENTF_RIGHTUP
MOUSEEVENTF_MIDDLEDOWN = win32con.MOUSEEVENTF_MIDDLEDOWN
MOUSEEVENTF_MIDDLEUP = win32con.MOUSEEVENTF_MIDDLEUP
KEYEVENTF_KEYUP = win32con.KEYEVENTF_KEYUP
MOUSEEVENTF_MOVE = win32con.MOUSEEVENTF_MOVE

# decimal vendor and product values
dev = usb.core.find(idVendor=0x046d, idProduct=0xc408)

# first endpoint
interface = 0
endpoint = dev[0][(0,0)][0]

# claim the device
usb.util.claim_interface(dev, interface)

KeyStatus = (0,0,0,0,0)
LastKeyDown = [0,0,0,0,0]
LastKeyUp = [0,0,0,0,0]
WheelMode = False
WheelUsed = False

def unpack(data):
    key1 = bool(data[0] & 0x1)
    key2 = bool(data[0] & 0x2)
    key3 = bool(data[0] & 0x4)
    key4 = bool(data[0] & 0x8)
    key5 = bool(data[0] & 0x10)
    if data[1] > 127:
        x = data[1] - 256
    else:
        x = data[1]
    if data[2] > 127:
        y = data[2] - 256
    else:
        y = data[2]

    if ROTATION or ACCELERATE:
        if ROTATION > 0:
            temp_x = math.cos(ROTATION) * x - math.sin(ROTATION) * y
            y = math.sin(ROTATION) * x + math.cos(ROTATION) * y
            x = temp_x

        if ACCELERATE > 0:
            x = x * pow(abs(x), ACCELERATE)
            y = y * pow(abs(y), ACCELERATE)

        if WheelMode:
            x = -x
            y = -y * 2.0

        x = int(x)
        y = int(y)
        
    if LEFTHAND:
        return ((key2, key1, key3, key5, key4), (x, y))
    return ((key1, key2, key3, key4, key5), (x, y))

while True:
    try:
        data = dev.read(endpoint.bEndpointAddress,endpoint.wMaxPacketSize)
        message = unpack(data)

        if message[1] != (0, 0):
            if WheelMode:
                win32api.mouse_event(MOUSEEVENTF_WHEEL, 0, 0, message[1][1])
                win32api.mouse_event(MOUSEEVENTF_HWHEEL, 0, 0, message[1][0])
                WheelUsed = True
            else:
                win32api.mouse_event(MOUSEEVENTF_MOVE, message[1][0], message[1][1], 0)
                
        if message[0][0] != KeyStatus[0]:
            if WheelMode:
                if message[0][0]:
                    win32api.mouse_event(MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0)
                else:
                    win32api.mouse_event(MOUSEEVENTF_MIDDLEUP, 0, 0, 0)
                WheelUsed = True
            else:
                if message[0][0]:
                    win32api.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0)
                else:
                    win32api.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0)

        if message[0][1] != KeyStatus[1]:
            WheelMode = message[0][1]
            
            if message[0][1]:
                WheelUsed = False
            else:
                if WheelUsed == False:
                    win32api.mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0)
                    win32api.mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0)

        if message[0][2] != KeyStatus[2]:
            if message[0][3]:
                win32api.mouse_event(MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0)
            else:
                win32api.mouse_event(MOUSEEVENTF_MIDDLEUP, 0, 0, 0)
                
        if message[0][3] != KeyStatus[3]:
            if message[0][3]:
                if WheelMode:
                    win32api.keybd_event(0x22,0,0,0)
                    win32api.keybd_event(0x22,0,KEYEVENTF_KEYUP,0)
                    WheelUsed = True
                else:
                    if time.time() - LastKeyDown[3] < 0.5:
                        win32api.keybd_event(0x11,0,0,0)
                        win32api.keybd_event(0x57,0,0,0)
                        win32api.keybd_event(0x57,0,KEYEVENTF_KEYUP,0)
                        win32api.keybd_event(0x11,0,KEYEVENTF_KEYUP,0)
                    LastKeyDown[3] = time.time()
            else:
                if WheelMode:
                    WheelUsed = True
                else:
                    if time.time() - LastKeyUp[3] > 0.5:
                        if time.time() - LastKeyDown[3] < 0.5:
                            win32api.keybd_event(166,0,0,0)
                        else:
                            win32api.keybd_event(0x23,0,0,0)
                            win32api.keybd_event(0x23,0,KEYEVENTF_KEYUP,0)
                LastKeyUp[3] = time.time()
                    
        if message[0][4] != KeyStatus[4]:
            if message[0][4]:
                if WheelMode:
                    win32api.keybd_event(0x21,0,0,0)
                    win32api.keybd_event(0x21,0,KEYEVENTF_KEYUP,0)
                    WheelUsed = True
                else:
                    if time.time() - LastKeyDown[4] < 0.5:
                        #win32api.keybd_event(0x74,0,0,0)
                        #win32api.keybd_event(0x74,0,KEYEVENTF_KEYUP,0)
                        win32api.keybd_event(0x11,0,0,0)
                        win32api.keybd_event(0x54,0,0,0)
                        win32api.keybd_event(0x54,0,KEYEVENTF_KEYUP,0)
                        win32api.keybd_event(0x11,0,KEYEVENTF_KEYUP,0)
                    LastKeyDown[4] = time.time()
            else:
                if WheelMode:
                    WheelUsed = True
                else:
                    if time.time() - LastKeyDown[4] < 0.5:
                        win32api.keybd_event(167,0,0,0)
                    else:
                        win32api.keybd_event(0x24,0,0,0)
                        win32api.keybd_event(0x24,0,KEYEVENTF_KEYUP,0)
                        #user32 = windll.LoadLibrary('user32.dll')
                        #user32.LockWorkStation()
                    LastKeyUp[4] = time.time()
                
        KeyStatus = message[0]
    except usb.core.USBError as e:
        data = None
        if e.args == ('Operation timed out',):
            continue
        
# release the device
usb.util.release_interface(dev, interface)
