
from fileinput import filename
import os
import zipfile
from windows_toasts import WindowsToaster, ToastText1
import time
from pywinauto import taskbar
import win32gui, win32con
from infi.systray import SysTrayIcon
import shutil

directory3DPrint = r"G:\3D Print"
directoryDownloads = r"G:\Downloads"

def commandWindow(systray):
    win32gui.ShowWindow(hide , win32con.SW_SHOW)
    time.sleep(15)
    win32gui.ShowWindow(hide , win32con.SW_HIDE)

def lastFolder(systray):
    os.startfile(outDirectory)

hide = win32gui.GetForegroundWindow()
win32gui.ShowWindow(hide , win32con.SW_HIDE)
menu_options = (("extractRar", None, commandWindow),("lastFolder", None, lastFolder))
systray = SysTrayIcon("3dPrinter.ico", "extractRar", menu_options)
systray.start()

ext = ('.zip', '.rar')
num = {'1', '2', '3', '4', '5', '6', '7', '8', '9',''}

while True:
    for files in os.listdir(directory3DPrint):
        if files.endswith(ext):
            fullPath = directory3DPrint + "/" + files
            print("3DfullPath: " + fullPath)
            fileName = os.path.basename(fullPath)[:-4] #removes extension
            print("3DfileName: " + fileName)

            # removes ID from beginning of string
            idCounter = 0
            for i in range(0, len(fileName)):
                print(fileName[i])
                if fileName[i] != "_" and fileName[i].isnumeric == True:
                    print(fileName[i])
                    idCounter = idCounter + 1
                else:
                    break
            if idCounter != 0:
                print('x')
                fileName = fileName[idCounter + 1:]

            #extract
            outDirectory = directory3DPrint + r"\\" + fileName
            print("3Dextract " + fullPath + " to " + outDirectory)
            with zipfile.ZipFile(fullPath, 'r') as zip_ref:zip_ref.extractall(outDirectory)
            os.remove(fullPath)

        
            path = os.path.realpath(outDirectory)

            wintoaster = WindowsToaster('Python')
            newToast = ToastText1()
            newToast.SetBody('Extracted: ' + fileName)
            newToast.on_activated = lambda _:  os.startfile(path)
            wintoaster.show_toast(newToast)

            time.sleep(10)


    for files in os.listdir(directoryDownloads):
        if files.endswith(ext):
            fullPath = directoryDownloads + '\\'  + files
            print("download fullpath " + fullPath)
            fileName = os.path.basename(fullPath)[:-4] #removes extension

            outDirectory = directoryDownloads + '\\'  + fileName
            print("(" + fullPath + ") to (" + outDirectory + ")")
            with zipfile.ZipFile(fullPath, 'r') as zip_ref:zip_ref.extractall(outDirectory)

            path = os.path.realpath(outDirectory)

            wintoaster = WindowsToaster('Python')
            newToast = ToastText1()
            newToast.SetBody('Extracted: ' + fileName)
            newToast.on_activated = lambda _:  os.startfile(path)
            wintoaster.show_toast(newToast)
            os.remove(fullPath)
            time.sleep(10)
