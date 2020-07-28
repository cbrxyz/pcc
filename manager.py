import time
import re
import winsound
import logging
from win10toast import ToastNotifier
from datetime import datetime

from cleaner import runCleaner
from windows import runWindows
from chrome import runChrome
from edge import runEdge

from helper import getVariable, warn, success, accentuate, log, VERSION, AUTHOR, REPO
from sheets import getVariables, sendResults, sendLogs, getLastRunDate, getMonthlyVals

logging.basicConfig(format='%(asctime)s - %(message)s', datefmt='%d-%b-%y %H:%M:%S')
logging.debug("hi")

def showToast(title, message, duration=3, threaded=True):
    toaster = ToastNotifier()
    toaster.show_toast(title,message,duration=duration,threaded=threaded)

def alertSound(sec):
    for _ in range(sec):
        winsound.Beep(800, 500)
        winsound.Beep(1000, 500)
    
def checkIfDate(api_key, monthly_reporter_debug, comp_name, from_email, to_email, subject):
    log("Checking if monthly logs need to be sent.")
    now = datetime.now()
    day = now.day
    lastDate = getLastRunDate(comp_name)
    if (lastDate == 1 and day == 1) or monthly_reporter_debug:
        print("Appears monthly logs need to be sent.")
        print("Preparing to send monthly logs...")
        getMonthlyVals(api_key, from_email, to_email, subject)

def run():
    accentuate("Starting the Public Computer Cleaner!")
    log(f"Created by {AUTHOR}. Version {VERSION}. Repo: {REPO}.\n")
    progVars = getVariables()
    if not progVars['on-switch']:
        warn("Stopping program. Global ON switch is set to 'OFF'.")
        return
    if progVars['monthly-reporter'] == progVars['computer-name']:
        checkIfDate(progVars['sendgrid-api-key'], progVars['monthly-reporter-debug'], progVars['computer-name'], progVars['from-email'], progVars['to-email'], progVars['email-subject'])
    else:
        log("Did not check if monthly report needed to be sent because this computer is not the monthly reporter.")
    runSecPref = progVars['alert-before-run']
    if runSecPref:
        log("Alerting before run, in accordance with variable setting.")
        showToast("Computer cleaner begins in 15 seconds.", "This removes all local information. To avoid this, please move the mouse.", duration=10, threaded=True)
        alertSound(1)
        time.sleep(9)
        showToast("Computer cleaner begins in 5 seconds.", "This removes all local information. To avoid this, please move the mouse.", duration=5, threaded=True)
        alertSound(5)
    log("Starting timer.")
    startCleanTime = time.perf_counter()
    log("Calling Cleaner algorithm...")
    cleanerResults = runCleaner()
    success("Completed Cleaner algorithm.")
    log("Calling Chrome algorithm...")
    chromeResults = runChrome()
    log("Calling Edge algorithm...")
    edgeResults = runEdge()
    log("Calling Windows algorithm...")
    windowsResults = runWindows(progVars['keep-name-regex'], progVars['keep-vendor-regex'])
    log("Completed all algorithms.")
    log("Compiling stats of this run.")
    res = {}
    res.update(cleanerResults)
    res.update(chromeResults)
    res.update(edgeResults)
    res.update(windowsResults)
    res.update({'times-ran': 1})
    log("Attempting to send results of run.")
    sendResults(progVars['computer-name'], res)
    sendLogs(progVars['computer-name'], res)
    endCleanTime = time.perf_counter()
    log(f'Completed in {endCleanTime - startCleanTime:0.4f} s')

run()