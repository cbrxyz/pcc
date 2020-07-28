import os

from helper import execute, accentuate, log
from cleaner import deleteAllWrapper

# Remember to make sure that your Edge version is 79 or greater!
edgeDefaultStr = r'C:\Users\{YOUR USERNAME HERE}\AppData\Local\Microsoft\Edge\User Data\Default'

def closeEdge():
    log("Attempting to close Edge.")
    instancesKilled = execute('taskkill /f /im msedge.exe', r'terminated')
    log("Killed " + str(instancesKilled) + " running instances of Edge.")
    return instancesKilled

def deleteEdgeHistory():
    results = deleteAllWrapper(edgeDefaultStr)
    return results

def runEdge():
    accentuate("Starting Edge Algorithm")
    instances = closeEdge()
    results = deleteEdgeHistory()
    log("Completed the Edge algorithm.")
    return {
        'edge-files-deleted': results[0],
        'edge-folders-deleted': results[1],
        'edge-instances-killed': instances
    }