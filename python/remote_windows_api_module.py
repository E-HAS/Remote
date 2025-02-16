import win32gui  # process in windows
import pyautogui 
import numpy as np
import cv2 # opencv capture
import win32api # mouse, keyboard
import win32con # mouse click 
import win32ui

from ctypes import windll
from PIL import Image

SysClass = ['ComboLBox','DDEMLEvent', '메시지', '#32768', '#32769', '#32770', '#32771', '#32772']
def initHwnds():
    hwnds = dict()
    
    def isEnumHandler(hwnd, output):
        isEnabled = win32gui.IsWindowEnabled(hwnd)
        isVisible = win32gui.IsWindowVisible(hwnd)

        isTitle = win32gui.GetWindowText(hwnd)
        isClass = win32gui.GetClassName(hwnd)

        isParent = str(win32gui.GetParent(hwnd))
        strHwnd = str(hwnd)
        if isEnabled and isVisible and (isTitle or isClass in SysClass):
            if isParent not in hwnds:
                hwnds[isParent] = dict()

            isHwnd = hwnds[isParent][strHwnd]=dict()
            isHwnd["title"]=isTitle
            isHwnd["class"]=isClass
        return True

    win32gui.EnumWindows(isEnumHandler, hwnds)
    return hwnds

def getParentHwnd(_hwnd):
    return str(win32gui.GetParent(_hwnd))
def getEnabledAndVisibleByHwnd(_hwnd):
    isEnabled = win32gui.IsWindowEnabled(_hwnd)
    isVisible = win32gui.IsWindowVisible(_hwnd)
    return isEnabled and isVisible
    

def getScreenShotDCByHwnd(_hwnd):
    try:
        _hwnd = int(_hwnd)
        left, top, right, bot = win32gui.GetWindowRect(_hwnd)
        w = right - left
        h = bot - top

        hwndDC = win32gui.GetWindowDC(_hwnd)
        mfcDC  = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()

        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)
        saveDC.SelectObject(saveBitMap)
        # 0 : False, 1: True
        result = windll.user32.PrintWindow(_hwnd, saveDC.GetSafeHdc(), 2)

        bmpinfo = saveBitMap.GetInfo()
        bmpstr = saveBitMap.GetBitmapBits(True)
        img = Image.frombuffer(
            'RGB',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRX', 0, 1)
        win32gui.DeleteObject(saveBitMap.GetHandle())
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(_hwnd, hwndDC)
        if result == True:
            np_image = np.array(img)
            frame = cv2.cvtColor(np_image, cv2.COLOR_RGB2BGR)
            f_width = frame.shape[1]
            f_height = frame.shape[0]

            strHwnd = str(_hwnd)
            if 'windowSize' not in encodedHwnds[strHwnd]:
                encodedHwnds[strHwnd]["windowsInfo"]=dict()
        
            encodedHwnds[strHwnd]["windowsInfo"]["left"]=left
            encodedHwnds[strHwnd]["windowsInfo"]["top"]=top
            encodedHwnds[strHwnd]["windowsInfo"]["right"]=right
            encodedHwnds[strHwnd]["windowsInfo"]["bot"]=bot
            encodedHwnds[strHwnd]["windowsInfo"]["width"]=f_width
            encodedHwnds[strHwnd]["windowsInfo"]["height"]=f_height

            return frame
        else:
            return None
    except Exception as e:
        print(f'getScreenShotDCByHwnd error : {e}')
        return None
    

def getScreenShotFrameByHwnd(_hwnd):
    left, top, right, bot = win32gui.GetWindowRect(_hwnd)
    w = right - left
    h = bot - top

    regionScreenShot = (left,top,w,h)
    frame = cv2.cvtColor(np.asarray(pyautogui.screenshot(region= regionScreenShot )), cv2.COLOR_RGB2BGR)
    f_width = frame.shape[1]
    f_height = frame.shape[0]

    if 'windowSize' not in encodedHwnds[_hwnd]:
        encodedHwnds[_hwnd]["windowsInfo"]=dict()
        
    encodedHwnds[_hwnd]["windowsInfo"]["left"]=left
    encodedHwnds[_hwnd]["windowsInfo"]["top"]=top
    encodedHwnds[_hwnd]["windowsInfo"]["right"]=right
    encodedHwnds[_hwnd]["windowsInfo"]["bot"]=bot
    encodedHwnds[_hwnd]["windowsInfo"]["width"]=f_width
    encodedHwnds[_hwnd]["windowsInfo"]["height"]=f_height

    return frame

def resizeScreenShotFrame(_hwnd, _frame):
    isWidth = encodedHwnds[_hwnd]["windowsInfo"]["width"];
    isHeight = encodedHwnds[_hwnd]["windowsInfo"]["height"];
    
    w_scale = 1.0
    h_scale = 1.0

    re_width = isWidth
    re_height = isHeight

    #기준치보다 클때 
    if SET_WIDTH < isWidth: 
        w_scale = float(isWidth/SET_WIDTH) #기준치 차이 비율
        re_width = SET_WIDTH
    if SET_HEIGHT < isHeight:
        h_scale = float(isHeight/SET_HEIGHT)
        re_height = SET_HEIGHT
    encodedHwnds[_hwnd]["windowsInfo"]["wScale"]=w_scale
    encodedHwnds[_hwnd]["windowsInfo"]["hScale"]=h_scale
    
    frame = cv2.resize(_frame, dsize=(re_width,re_height), interpolation=cv2.INTER_AREA)
    f_width = frame.shape[1]
    f_height = frame.shape[0]
    encodedHwnds[_hwnd]["windowsInfo"]["rWidth"]=f_width
    encodedHwnds[_hwnd]["windowsInfo"]["rHeight"]=f_height
    
    return frame
    
encodedHwnds=dict()
currentRemoteHwnd=str()

SET_WIDTH = 960
SET_HEIGHT = 600
def getByteFromScreenShotFrameByHwnd(_hwnd):
    if _hwnd is not None:
        global currentRemoteHwnd
        
        if _hwnd not in encodedHwnds:
            encodedHwnds[_hwnd]=dict()

        if  currentRemoteHwnd != _hwnd:
            win32gui.ShowWindow(_hwnd,5) #윈도우 보이도록하기
            #win32gui.ShowWindow(_hwnd,3) #최대화 하기

            if currentRemoteHwnd != '':
                win32gui.ShowWindow(currentRemoteHwnd,11)#윈도우 최소화하기            
        currentRemoteHwnd=_hwnd

        frame = getScreenShotFrameByHwnd(_hwnd)
        frame = resizeScreenShotFrame(_hwnd, frame)

        _ , encoded_image = cv2.imencode('.png',frame)
        
        return encoded_image.tobytes(), encodedHwnds[_hwnd]["windowsInfo"]
    else:
        return None
def getByteFromScreenShotDCByHwnd(_hwnd):
    if _hwnd is not None:
        global currentRemoteHwnd
        
        if _hwnd not in encodedHwnds:
            encodedHwnds[_hwnd]=dict()

        if  currentRemoteHwnd != _hwnd:
            win32gui.ShowWindow(_hwnd,5) #윈도우 보이도록하기
            #win32gui.ShowWindow(_hwnd,3) #최대화 하기

            if currentRemoteHwnd != '':
                win32gui.ShowWindow(currentRemoteHwnd,11)#윈도우 최소화하기            
        currentRemoteHwnd=_hwnd

        frame = getScreenShotDCByHwnd(_hwnd)
        frame = resizeScreenShotFrame(_hwnd, frame)

        _ , encoded_image = cv2.imencode('.png',frame)
        
        return encoded_image.tobytes(), encodedHwnds[_hwnd]["windowsInfo"]
    else:
        return None

def movingMouseInWindowHwnd(_hwnd,_event, _x,_y):
    global encodedHwnds
    encodedSize = encodedHwnds[_hwnd]["windowsInfo"]

    pos=( int(float(encodedSize["left"])+float(_x)) , int(float(encodedSize["top"])+float(_y)) )
    win32api.SetCursorPos(pos)
    if _event == 'LClick':
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, pos[0], pos[1], 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, pos[0], pos[1], 0, 0)
    


def pressedKeyboardInWindowHwnd(_hwnd, _key):
    pyautogui.write(_key)
    
    
