import time
import os
import subprocess
import re

#Files
from cleaner import deleteAllFiles
from helper import execute, accentuate, log

chromeDataStr = r'PLACE YOUR CHROME DATA STRING HERE'

def closeChrome():
    log("Attempting to close Chrome.")
    instancesKilled = execute('taskkill /f /im chrome.exe', r'terminated')
    log("Killed " + str(instancesKilled) + " running instances of Chrome.")
    return instancesKilled

def deleteHistory():
    filesDeleted = deleteAllFiles(chromeDataStr)
    log("Deleted " + str(filesDeleted) + " Chrome files.")
    numFolderDeleted = execute(r'FOR /D %p IN ("' + chromeDataStr + r'\*.*") DO (if not "%~np"=="Default" rmdir "%p" /s /q)', r'rmdir') - 1
    log("Removed " + str(numFolderDeleted) + " Chrome storage folders.")
    return [filesDeleted, numFolderDeleted]

def runChrome():
    accentuate("Starting Chrome Algorithm")
    instancesKilled = closeChrome()
    chromeResults = deleteHistory()
    log("Completed the Chrome algorithm.")
    return {
        'chrome-instances-killed': instancesKilled,
        'chrome-files-deleted': chromeResults[0],
        'chrome-folders-deleted': chromeResults[1]
    }