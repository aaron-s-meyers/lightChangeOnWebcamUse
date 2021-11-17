import win32api
import win32event
import win32con
import winreg
import threading
import requests
import json
import time


#hue stuff
# Credentials stored in JSON file
f = open('cred.json',)
cred = json.load(f)
f.close()
headers = {
    'Accept': 'application/json',
}
on = '{"on":true,"bri":255,"sat":255,"hue":322}'
off = '{"on":false,"bri":255,"sat":255,"hue":0}'
ip = cred["ip"]
username = cred["username"]
hueLight = 6


def notify(path, type):
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,path,0,winreg.KEY_READ)
    evt = win32event.CreateEvent(None, 0, 0, None)
    win32api.RegNotifyChangeKeyValue(key, 1, win32api.REG_NOTIFY_CHANGE_LAST_SET, evt, True)
    ret_code=win32event.WaitForSingleObject(evt,-1)
    if ret_code == win32con.WAIT_OBJECT_0:
        time.sleep(2)
        currKey = key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,path,0,winreg.KEY_READ)
        if (winreg.QueryValueEx(currKey, "LastUsedTimeStop")[0] == 0):  #last used value is 0 when webcam is in use
            print(type + " open")
            #turn lroom light red
            requests.put(f'http://{ip}/api/{username}/lights/{hueLight}/state', headers=headers, data=on)
        else:
            print (type + " closed")
            #turn lroom light off
            requests.put(f'http://{ip}/api/{username}/lights/{hueLight}/state', headers=headers, data=off)
        if (type == "zoom"):
            zoom()   ##restart monitor
        elif (type == "nexi"):
            nexi()   ##restart monitor
        elif (type == "chrome"):
            chrome() ##restart monitor
    if ret_code == win32con.WAIT_TIMEOUT:
        print ("TIMED")

def thread(path, type):
    threading.Thread(target=notify,
            args=(path, type,),
        ).start()

def zoom():
    thread(r'SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\webcam\NonPackaged\C:#Users#aaron#AppData#Roaming#Zoom#bin#Zoom.exe', 'zoom',)

def nexi():
    thread(r'SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\webcam\NonPackaged\C:#Users#aaron#AppData#Local#Nexi#Nexi.exe', 'nexi',)

def chrome():
    thread(r'SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\webcam\NonPackaged\C:#Program Files (x86)#Google#Chrome#Application#chrome.exe', 'chrome',)


#first run
zoom()
nexi()
chrome()