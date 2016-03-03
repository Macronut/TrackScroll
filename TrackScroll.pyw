import pythoncom
import pyHook
import time
import win32api
import win32con

MousePosition = (0, 0)
WheelMode = False
WheelModeTime = 0
WheelValue = 0
CloseMode = False
CloseTime = 0

def OnMouseEvent(event):
    global WheelMode, MousePosition, WheelModeTime, WheelValue
    
    if WheelMode:
        if event.Message == 512:
            pos_x = int(event.Position[0])
            pos_y = int(event.Position[1])
            wheelvalue = int(MousePosition[1]) - pos_y
            wheelvalue *= 2
            if event.WindowName == 'Chrome Legacy Window':
                win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, wheelvalue)
            else:
                WheelValue += wheelvalue
                if (WheelValue >= 40 or WheelValue <= -40):
                    win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, WheelValue)
                    WheelValue = 0
            return False
    else:
        MousePosition = event.Position
        
    return True

def OnKeyboardEvent(event):
    global WheelMode, WheelModeTime, CloseMode, CloseTime
    
    if event.KeyID == 145:
        if event.Message == 256:
            WheelMode = True
            WheelModeTime = time.time()
        else:
            if (time.time() - WheelModeTime < 0.5):
                win32api.keybd_event(166,0,0,0)
            WheelMode = False
        return False
    elif event.KeyID == 19:
        if event.Message == 256:
            if CloseMode and (time.time() - CloseTime < 0.6):
                win32api.keybd_event(17,0,0,0)
                win32api.keybd_event(87,0,0,0)
                win32api.keybd_event(87,0,win32con.KEYEVENTF_KEYUP,0)
                win32api.keybd_event(17,0,win32con.KEYEVENTF_KEYUP,0)
                CloseMode = False
            else:
                CloseMode = True
            CloseTime = time.time()
        else:
            if (time.time() - CloseTime < 0.5):
                win32api.keybd_event(167,0,0,0)
                
        return False
    return True

hm = pyHook.HookManager()
hm.MouseAll = OnMouseEvent
hm.KeyAll = OnKeyboardEvent
hm.HookMouse()
hm.HookKeyboard()
pythoncom.PumpMessages()
