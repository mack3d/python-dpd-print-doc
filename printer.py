import time
import win32api
import win32print

def printFile(file, printerName = '2727'):
    defaultPrinter = win32print.GetDefaultPrinter()
    if defaultPrinter != printerName:
        win32print.SetDefaultPrinter(printerName)
    win32api.ShellExecute(0, "print", file, None,  ".",  0)
    time.sleep(5)
    win32print.SetDefaultPrinter(defaultPrinter)