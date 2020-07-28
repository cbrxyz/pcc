# pylint: skip-file
# Used with odd win32com import error

import ctypes
import wmi
import pywinauto
import re
import win32com.shell.shell as shell

from helper import accentuate, warn, log

imgStr = r'C:\Backgrounds\taxis.jpg'

def setBackground():
    ctypes.windll.user32.SystemParametersInfoW(20, 0, imgStr, 0)

def uninstallUselessPrograms(nameString, vendorString):
    w = wmi.WMI()
    deletedNames = []
    for p in w.Win32_Product():
        nameMatches = re.findall(nameString, p.name, re.I)
        vendorMathces = re.findall(vendorString, p.vendor, re.I)
        if len(nameMatches) > 0: 
            log("Keeping " + p.name + " because it matched a name keep RegEx.")
        if len(vendorMathces) > 0: 
            log("Keeping " + p.name + " because it matched a vendor keep RegEx.")
        else: 
            warn("Deleting " + p.name + ".")
            deleteProgram(p.name)
            log("Attempted to delete " + p.name + ".")
            deletedNames.append(p.name)
    allNames = []
    for p in w.Win32_Product():
        allNames.append(p.name)
    for d in deletedNames:
        if d not in allNames:
            log("Confirmed successful deletion of " + d + ".")
    return len(deletedNames)

def deleteProgram(name):
    commands = "wmic product where name='" + name + "' call uninstall"
    shell.ShellExecuteEx(lpVerb='runas', lpFile='cmd.exe', lpParameters='/c '+commands)

def runWindows(nameString, vendorString):
    accentuate("Starting Windows Algorithm")
    setBackground()
    log("Set computer background.")
    numDeleted = uninstallUselessPrograms(nameString, vendorString)
    log("Completed Windows algorithm.")
    return {
        'programs-uninstalled': numDeleted
    }