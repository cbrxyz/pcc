import os
import re

from helper import execute, accentuate, log

documentsStr = r"DOCUMENTS PATH GOES HERE"
desktopStr = r"DESKTOP PATH GOES HERE"
picturesStr = r"PICTURES PATH GOES HERE"
videosStr = r"VIDEOS PATH GOES HERE"
musicStr = r"MUSIC PATH GOES HERE"
downloadsStr = r"DOWNLOADS PATH GOES HERE"

def deleteAllDesktopShortcuts():
    output = execute('for /R ' + desktopStr + ' %f in (*) do (if not "%~xf"==".lnk" del "%~f")', '', True)
    tests = re.findall(r'== "\.lnk"', output)
    nonDeletions = re.findall(r'"\.lnk" == "\.lnk"', output)
    log("Deleted " + str(len(tests) - len(nonDeletions)) + " files from the Desktop.")
    return len(tests) - len(nonDeletions)

def deleteAll(directory):
    filesDeleted = deleteAllFiles(directory)
    foldersDeleted = deleteAllFolders(directory)
    return [filesDeleted, foldersDeleted]

def deleteAllFiles(directory):
    return execute('del /S /Q "' + directory + '"', r'Deleted file')

def deleteAllFolders(directory):
    return execute('FOR /D %p IN ("' + directory + r'\*.*") DO rmdir "%p" /s /q', r'rmdir')

def deleteAllWrapper(directory):
    outputs = deleteAll(directory)
    log("Deleted " + str(outputs[0]) + " files and " + str(outputs[1]) + " folders from " + directory + ".")
    return outputs

def runCleaner():
    accentuate("Starting Cleaner Algorithm")
    numDesktopFiles = deleteAllDesktopShortcuts()
    desktopFoldersDeleted = deleteAllFolders(desktopStr)
    log("Deleted " + str(desktopFoldersDeleted) + " folders from the Desktop.")
    documentsDeleted = deleteAllWrapper(documentsStr)
    picturesDeleted = deleteAllWrapper(picturesStr)
    videosDeleted = deleteAllWrapper(videosStr)
    musicDeleted = deleteAllWrapper(musicStr)
    downloadsDeleted = deleteAllWrapper(downloadsStr)
    log("Completed the cleaner algorithm.")
    return {
        'desktop-files-deleted': numDesktopFiles,
        'desktop-folders-deleted': desktopFoldersDeleted,
        'document-files-deleted': documentsDeleted[0],
        'document-folders-deleted': documentsDeleted[1],
        'pictures-files-deleted': picturesDeleted[0],
        'pictures-folders-deleted': picturesDeleted[1],
        'videos-files-deleted': videosDeleted[0],
        'videos-folders-deleted': videosDeleted[1],
        'music-files-deleted': musicDeleted[0],
        'music-folders-deleted': musicDeleted[1],
        'downloads-files-deleted': downloadsDeleted[0],
        'downloads-folders-deleted': downloadsDeleted[1],
    }