import sys
import os
import os.path
import time
import win32gui as gui
import webbrowser as wb
import json
import datetime
import smtplib
from collections import Counter
import pynput
from pynput import mouse, keyboard
from pynput.keyboard import Key, Listener
from pynput.mouse import Controller
import keyboard as kb

#establish global vars
keysPerMinute = []
mouseX = []
mouseY = []
xWeight = []
yWeight = []
day = ''
cur_time = ''
active_window = ''
cur_window = ''
backspaces = 0
flag = 0
studyTime = 10000 #the larger this is, the more it will study before enforcing
state = 'study'



#START UP FOLDER GENERATION, ONLY RUNS IF THE FOLDERS AREN'T THERE
def GenerateFolders():
    #create initial user json
    userJson = {}
    userJson['loopData']=[]
    userJson['loopData'].append({
                        "active_windows": '',
                        "x_coordinates": [],
                        "y_coordinates": [],
                        "x_weights": [],
                        "y_weights": [],
                        "backspaces": [],
                        "keystrokes_per_loop": [],
                        "day": [],
                        "time": []
                    }
                    )
    Json_Report_Util(userJson,'userdata\\USER.txt')

    #create the meta json file
    name = os.getlogin()
    email = input("Recovery Email: ")
    metaJson = {}
    metaJson['User']=[]
    metaJson['User'].append({
                        "name": name,
                        "email": email 
                    }
                    )
    Json_Report_Util(metaJson,'userdata\\META.txt')      

    SendReport('userdata\\','p_counter.txt','0')        
#CHECKS TO SEE IF NECCESSARY FOLDERS ARE THERE
def isSetup():
    path = "D:\\School\\FINAL SEMESTER\\Senior Project\\userdata\\META.txt"
    if (os.path.isfile(path)== False):
        GenerateFolders()

# UTILITY FUNCTIONS {

#tells the system when it's been long enough to switch to the enforce state
def switchState():
    s = False
    global studyTime
    with open ('userdata\\p_counter.txt','r') as f:
        for i in f:
            temp = int(i)
            if (temp > studyTime):
                s= True
    return s

def Json_Report_Util(data, path):
    with open (path, "w") as f:
        json.dump(data,f, indent=4)
    
def SendReport(path,filename,data):
    filepath = os.path.join(path,filename)
    f = open(filepath,'a')
    f.write(data)
    f.close
    
#File used to keep track of how many times a report has been sent
def Increment():
    k=0
    with open ('userdata\\p_counter.txt','r') as f:
        for i in f:
            k = int(i)+1
    with open ('userdata\\p_counter.txt','w') as f:
        f.write(str(k))

def Reset():
    with open ('userdata\\p_counter.txt','w') as f:
        f.write('0')
        
def organizeData():
    global active_window, mouseX, mouseY, xWeight, yWeight, backspaces, keysPerMinute, day, cur_time  
    path = "userdata\\USER.txt"
    with open(path, "r+") as json_file:
        data = json.load(json_file)
        temp = data['loopData']
        info = {"active_windows": str(active_window),
                "x_coordinates": str(mouseX),
                "y_coordinates": str(mouseY),
                "x_weights": str(xWeight),
                "y_weights": str(yWeight),
                "backspaces": str(int(backspaces/100)),
                "keystrokes_per_loop": str(len(keysPerMinute)),
                "day": day,
                "time": cur_time}
        temp.append(info)
    Increment()
    Json_Report_Util(data,path)
    
def removeEmpty():
    path = "userdata\\USER.txt"
    with open(path, "r+") as json_file:
        data = json.load(json_file)
        for i in data['loopData']:
            if (i['time'] == ''):
                print('empty found')
                del i
#KEYS PER MINUTE BEHAVIOR
def on_press(key):
    global keysPerMinute
    keysPerMinute.append(key)

def on_release(key):
    return

def Beh_KeysPerMinute():
    global state
    listener = keyboard.Listener(
        on_press=on_press,
        on_release=on_release)
    listener.start()

    if (state == 'report'):
        listener.stop()
        return
    
#BACKSPACES PER MINUTE BEHAVIOR
def Beh_BackspaceCount():
    return kb.is_pressed('backspace')

    
#MOUSE MOVEMENT BEHAVIOR
    
def Beh_mouseMovement():
    global mouseX, mouseY
    pos =gui.GetCursorPos()
    mouseX.append(pos[0])
    mouseY.append(pos[1])
    
def mouseMovementCleanup():
    global mouseX, mouseY
    for i in range(len(mouseX)):
        mouseX[i] = int(mouseX[i]/10) 
        mouseY[i] = int(mouseY[i]/10)
    mouseWeights()

def mouseWeights():
    global mouseX, mouseY, xWeight, yWeight
    
    xWeight = Counter(mouseX)
    yWeight = Counter(mouseY)

        
#DATE AND TIME BEHAVIOR
def get_dateTime():
    global day,cur_time
    t = datetime.datetime.now()
    day = t.strftime("%A")
    cur_time = t.strftime("%H:%M:%S")

#ACTIVE WINDOW BEHAVIOR
def activeWindow():
    w = gui
    return str(w.GetWindowText(w.GetForegroundWindow()))

def windowCleanup():
    global active_window
    active_window = list(dict.fromkeys(active_window))
    tmp = ''
    for i in active_window:
        tmp += i
    active_window = tmp    
#ENFORCEMENT FUNCTIONS

def compare_xWeight(r,xWeight):
    out = False
    f1 , f2 = str(r['x_weights']).find('{'), str(r['x_weights']).rfind('}')
    newR = str(r['x_weights'])[f1:f2+1]
    r['x_weights'] = newR
    tempList = str(r['x_weights']).split(',')
    for e in tempList:
        vI = e[(e.find('{')+len('{')):e.find(':')]
        vW = e[(e.find(':')+len(':')):e.find('}')]

        for i in xWeight:
            if (vI == i):
                if (vW== xWeight[i]):
                    out = True
    return out

def compare_yWeight(r,yWeight):
    out = False
    f1 , f2 = str(r['y_weights']).find('{'), str(r['y_weights']).rfind('}')
    newR = str(r['y_weights'])[f1:f2+1]
    r['y_weights'] = newR
    tempList = str(r['y_weights']).split(',')
    for e in tempList:
        vI = e[(e.find('{')+len('{')):e.find(':')]
        print(vI)
        vW = e[(e.find(':')+len(':')):e.find('}')]

        for i in yWeight:
            if (vI == i):
                if (vW== yWeight[i]):
                    out = True
    return out

def compare_back(r,bs):
    out = False
    wiggle = int(bs)/10
    out = (abs(int(r['backspaces']) - float(bs))< wiggle)
    return out

def compare(r):
    out = False
    global xWeight, yWeight, backspaces
    if (compare_xWeight(r,xWeight) and compare_yWeight(r,yWeight)):
        out = True
    if (compare_back(r,backspaces)):
        out = True
    return out
    

def compareData():
    out = False
    active_window = ''
    path = "userdata\\USER.txt"
    with open(path, "r") as json_file:
        data = json.load(json_file)
        temp = data['loopData']
        for r in temp:
            if (r['active_windows'] == active_window):
                out = compare(r)
                if (out):
                    return out

    return out

#FLAG FUNCTIONS
def getEmail():
    e = ''
    path = "userdata\\META.txt"
    with open(path, "r") as json_file:
        data= json.load(json_file)
        e = data['User']['email']
        print(e)
    return e

def getReport():
    global active_window, mouseX, mouseY, xWeight, yWeight, backspaces, keysPerMinute, day, cur_time
    info = {"active_windows": str(active_window),
                "x_coordinates": str(mouseX),
                "y_coordinates": str(mouseY),
                "x_weights": str(xWeight),
                "y_weights": str(yWeight),
                "backspaces": str(backspaces/100),
                "keystrokes_per_loop": str(len(keysPerMinute)),
                "day": day,
                "time": cur_time}
    s = ('Abnormal Behavior detected at your work station: \n' + str(info))
    return s
    
#MAIN LOOP
def main():
    isSetup()
    #establish variables 
    global state 
    global keysPerMinute
    global mouseX, mouseY
    global backspaces
    global active_window, cur_window
    global flag
    keyLock = False
    resetTime = 30
    loopStartTime = 0
    tick = 1
    counter = 0
    #time loop
    
    while(True):
        while(state == 'enforce'):
            alert = False
            get_dateTime()
            mouseMovementCleanup()
            windowCleanup()
            removeEmpty()
            #run comparison function
            if (compareData()== False):
                flag += 1
                if (flag >3):
                    alert = True
            else:
                flag = 0
            print('comparing data')
            loopStartTime = tick
            backspaces = 0 
            counter = 0
            keyLock = False
            mouseX.clear()
            mouseY.clear()
            xWeight.clear()
            yWeight.clear()
            active_window = cur_window
            day = ''
            cur_time = ''
            keysPerMinute.clear()
            #comparison approved
            if (alert == False):
                state = 'study'
            else:
                state = 'flag'

        while(state == 'report'):
            #send reports
            get_dateTime()
            mouseMovementCleanup()
            windowCleanup()
            organizeData()

            #restart loop
            print("report sent")
            print(active_window)
            #Reset Vars
            loopStartTime = tick
            backspaces = 0 
            counter = 0
            keyLock = False
            mouseX.clear()
            mouseY.clear()
            xWeight.clear()
            yWeight.clear()
            active_window = cur_window
            day = ''
            cur_time = ''
            keysPerMinute.clear()

            if (switchState()):
                removeEmpty()
                state = 'enforce'
            else:
                state = 'study'

        while(state == 'flag'):
            #send email
            server = smtplib.SMTP_SSL('smtp.mail.yahoo.com', 465)

            server.login("njgrimaldi@yahoo.com", "m3CD8M1jSe3PN8")
            server.sendmail("njgrimaldi@yahoo.com",
                            getEmail(),
                            getReport())
            server.quit()
            #reset vars and incrementer
            Reset()
            state = 'study'
            #logout
            os.system("shutdown -1")

        while(state == 'study'):
            #main loop           
            #check for backspace, if true increment
            if (Beh_BackspaceCount()):
                backspaces +=1
            #run keysPerMinute, but only once per loop
            if (keyLock == False):
                keyLock = True
                Beh_KeysPerMinute()
            
            #display loop time for debug purposes and increment tick
            if (tick != int(time.time())):
                counter += 1
                counter = counter%30

                #run mouseMovement once per second
                Beh_mouseMovement()
            #change to report state after loop is complete
            cur_window = activeWindow()
            if (active_window != cur_window):
                if (switchState()):
                    state = 'enforce'
                else:
                    state = 'report'
            tick = int(time.time())
            

if __name__ == "__main__":
    #run main
    main()

    

