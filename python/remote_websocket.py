import json
import asyncio
import websockets
import base64
import time
import remote_windows_api_module as remoteWindowAPI
from enum import Enum

def getObjectByJson(_str): #Json > str
    return json.loads(_str)

def getJsonByObject(_str): #str > Json
    return json.dumps(_str)

class RequestType(Enum):
    PROCESSLIST ='0'
    SCREENSHOT = '1'
    MOVECURSOR = '2'
    
async def messageHandler(_websocket, _msg):
    request = _msg['request']
    reqType = _msg['reqType'] if 'reqType' in _msg else None
    hwnd = _msg['hwnd'] if 'hwnd' in _msg else None
    
    data=dict()
    if request == RequestType.PROCESSLIST.value:
        async with HWNDS_LOCK:
            global HWNDS
            
            data["ProcessList"]=[]
            for hwnd in HWNDS["0"]:
                title = HWNDS["0"][hwnd]['title']
                data["ProcessList"].append([hwnd, title])
                
    elif request == RequestType.SCREENSHOT.value:
        taskName = 'sendScreenShotBy'+str(hwnd)
        stopTaskByName(taskName)

        if reqType =='1':
            #sendScreenShotByHwnd(hwnd,_websocket)
            task = asyncio.create_task(sendScreenShotByHwnd(hwnd,_websocket))
            #task.set_name(taskName)
            
    elif request == RequestType.MOVECURSOR.value:
        if reqType == 'mouse':
            event = _msg['event']
            x = _msg['x']
            y = _msg['y']
            remoteWindowAPI.movingMouseInWindowHwnd(hwnd,event, x,y)
        elif reqType == 'keyboard':
            key = _msg['key']
            remoteWindowAPI.pressedKeyboardInWindowHwnd(hwnd, key)
        
            
    msg = dict()
    msg["message"]="ok"
    msg["request"]=request
    msg["data"]=data
    return msg

# websocket 1-3
async def handler(websocket):
    print('start handler')
    
    while True:
        try:
            preMessage = await websocket.recv()
            recvPreMessage = getObjectByJson(preMessage) #Dict형으로 변환
            
            postMessage = await messageHandler(websocket,recvPreMessage) #SendHandler 

            sendPostMessage = getJsonByObject(postMessage)#Json형으로 변환
            await websocket.send(sendPostMessage)#데이터 보내기
        except websockets.ConnectionClosedOK:
            break
        
HWNDS = dict()
HWNDS_LOCK = asyncio.Lock()
FRAME_RATE = float(1/60)
async def initHwnds():
    global HWNDS
    while True:
        async with HWNDS_LOCK:
            HWNDS.clear()
            HWNDS = remoteWindowAPI.initHwnds()
        await asyncio.sleep(1)
        
async def sendScreenShotByHwnd(_hwnd, _socket):
    status = remoteWindowAPI.getEnabledAndVisibleByHwnd(_hwnd)
    if status:
        encoded, windowsInfo = remoteWindowAPI.getByteFromScreenShotDCByHwnd(_hwnd)
        data = dict()
        data['hwnd'] = _hwnd

        data['posX'] = windowsInfo['left']
        data['posY'] = windowsInfo['top']
        data['wScale'] = windowsInfo['wScale']
        data['hScale'] = windowsInfo['hScale']
        data['width'] = windowsInfo['width']
        data['height'] = windowsInfo['height']
        data['encoded'] = base64.b64encode(encoded).decode()
            
        send = dict()
        send["message"]="ok"
        send["request"]=RequestType.SCREENSHOT.value
        send["data"]=data
            
        sendJsonMessage = getJsonByObject(send)
        await _socket.send(sendJsonMessage)
    else:
        return
    
async def sendAsyncScreenShotByHwnd(_hwnd, _socket):
    while True:
        status = remoteWindowAPI.getEnabledAndVisibleByHwnd(_hwnd)
        if status:
            encoded, windowsInfo = remoteWindowAPI.getByteFromScreenShotDCByHwnd(_hwnd)
            data = dict()
            data['hwnd'] = _hwnd

            data['posX'] = windowsInfo['left']
            data['posY'] = windowsInfo['top']
            data['wScale'] = windowsInfo['wScale']
            data['hScale'] = windowsInfo['hScale']
            data['width'] = windowsInfo['width']
            data['height'] = windowsInfo['height']
            data['encoded'] = base64.b64encode(encoded).decode()
            
            send = dict()
            send["message"]="ok"
            send["request"]=RequestType.SCREENSHOT.value
            send["data"]=data
            
            sendJsonMessage = getJsonByObject(send)
            await _socket.send(sendJsonMessage)
            await asyncio.sleep(FRAME_RATE)
        else:
            break

def stopTaskByName(_name):
    loop = asyncio.get_running_loop()
    pending = asyncio.all_tasks(loop=loop)
    for task in pending:
        taskName = task.get_name()
        if taskName == _name:
            task.cancel()
    
# websocket 1-2
async def main():
    async with websockets.serve(handler, "localhost",8000):
        task = asyncio.create_task(initHwnds())
        task.set_name('initHwnds')
        
        await asyncio.Future()

if __name__ == '__main__':
    # websocket 1-1
    asyncio.run(main(), debug=False)
    
