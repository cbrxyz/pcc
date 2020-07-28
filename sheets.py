import gspread
import time
import datetime
import pandas as pd
import colorama
import re

from helper import getVariable, warn, log
from emailer import sendMonthlyReport

spreadsheetName = r'SPREADSHEET NAME HERE'

gc = gspread.service_account(filename="service_account.json")

sh = gc.open(spreadsheetName)

def getVariables():
    log("Attempting to get values for needed variables.")
    vals = sh.worksheet("Global Variables").get('B5:J14')
    log("Obtained values from Google Sheet.")
    # SendGrid API Key
    sendGridAPIKey = getVariable('SENDGRID_API_KEY')
    if sendGridAPIKey == False:
        try:
            sendGridAPIKey = vals[0][2]
        except:
            raise ValueError("SENDGRID_API_KEY has no value.")
    # On/off switch
    onOff = getVariable('GLOBAL_ON_SWITCH')
    if onOff == False:
        try:
            onOff = vals[1][2]
            if len(onOff) < 1:
                warn("GLOBAL_ON_SWITCH not formally declared. Assuming ON.")
                onOff = True
            else:
                onOff = (onOff == "ON")
        except:
            warn("GLOBAL_ON_SWITCH not formally declared. Assuming ON.")
            onOff = True
    else:
        onOff = (onOff == "ON")
    # Alert Before Run option
    alertBeforeRun = getVariable('ALERT_BEFORE_RUN')
    if alertBeforeRun == False:
        try:
            alertBeforeRun = vals[2][2]
            if len(alertBeforeRun) < 1:
                warn("ALERT_BEFORE_RUN not formally declared. Assuming ON.")
                alertBeforeRun = True
            else:
                alertBeforeRun = (alertBeforeRun == "ON")
        except:
            warn("ALERT_BEFORE_RUN not formally declared. Assuming ON.")
            alertBeforeRun = True
    else:
        alertBeforeRun = (alertBeforeRun == "ON")
    # Keep Name RegEx
    keepNameRegex = getVariable('KEEP_NAME_REGEX')
    if keepNameRegex == False:
        try:
            keepNameRegex = vals[3][2]
        except:
            raise ValueError("KEEP_NAME_REGEX has no value.")
    # Keep Vendor RegEx
    keepVendorRegex = getVariable('KEEP_VENDOR_REGEX')
    if keepVendorRegex == False:
        try:
            keepVendorRegex = vals[4][2]
        except:
            raise ValueError("KEEP_VENDOR_REGEX has no value.")
    # Computer Name
    compName = getVariable('COMPUTER_NAME')
    if compName == False:
        raise ValueError("COMPUTER_NAME is not defined in settings.txt. Without this, the program can not identify which page to send logs to.")
    # Monthly Reporter
    monthlyReporter = getVariable('MONTHLY_REPORTER')
    if monthlyReporter == False:
        try:
            monthlyReporter = vals[5][2]
        except:
            raise ValueError("MONTHLY_REPORTER has no value.")
    # Monthly Reporter Debug
    monthlyReporterDebug = getVariable('MONTHLY_REPORTER_DEBUG')
    if monthlyReporterDebug == False:
        try:
            monthlyReporterDebug = vals[6][2]
            if len(monthlyReporterDebug) < 1:
                warn("MONTHLY_REPORTER_DEBUG not formally declared. Assuming OFF.")
                monthlyReporterDebug = False
            else:
                monthlyReporterDebug = (monthlyReporterDebug == "ON")
        except:
            warn("MONTHLY_REPORTER_DEBUG not formally declared. Assuming OFF.")
            monthlyReporterDebug = False
    else:
        monthlyReporterDebug = (monthlyReporterDebug == "ON")
    # From Email
    fromEmail = getVariable('FROM_EMAIL')
    if fromEmail == False:
        try:
            fromEmail = vals[7][2]
        except:
            raise ValueError("FROM_EMAIL has no value.")
    # To Email
    toEmail = getVariable('TO_EMAIL')
    if toEmail == False:
        try:
            toEmail = vals[8][2]
        except:
            raise ValueError("TO_EMAIL has no value.")
    # Subject
    emailSubject = getVariable('EMAIL_SUBJECT')
    if emailSubject == False:
        try:
            emailSubject = vals[9][2]
            if len(emailSubject) < 1:
                warn("EMAIL_SUBJECT not formally declared. Assuming ' [AUTO] Public Computer Cleaner Monthly Report'.")
                emailSubject = ' [AUTO] Public Computer Cleaner Monthly Report'
        except:
            warn("EMAIL_SUBJECT not formally declared. Assuming OFF.")
            emailSubject = ' [AUTO] Public Computer Cleaner Monthly Report'
    log("Successfully attemped to find all values for program.")
    return {
        'sendgrid-api-key': sendGridAPIKey,
        'on-switch': onOff,
        'alert-before-run': alertBeforeRun,
        'keep-name-regex': keepNameRegex,
        'keep-vendor-regex': keepVendorRegex,
        'computer-name': compName,
        'monthly-reporter': monthlyReporter,
        'monthly-reporter-debug': monthlyReporterDebug,
        'from-email': fromEmail,
        'to-email': toEmail,
        'email-subject': emailSubject,
    }

def getLastRunDate(comp_name):
    cell = sh.worksheet(comp_name).acell("B6").value
    matches = re.match(r"([0-9]*)", cell)
    if len(matches[1]) > 1:
        days = int(matches[1])
    else:
        days = 0
    return days

def clearMonthlyStats(computerName):
    sh.worksheet(computerName).update('H9:I16', [
        [0, 0],
        [0, 0],
        [0, 0],
        [0, 0],
        [0, 0],
        [0, 0],
        [0, 0],
        [0, 0],
    ])
    sh.worksheet(computerName).update('J20:J23', [
        [0],
        [0],
        [0],
        [0],
    ])
    log(f"Cleared monthly statistics for {computerName}.")

def sendLogs(computerName, r):
    lastRow = sh.worksheet("Log - " + computerName).acell('V1').value
    lastRow = int(lastRow)
    mostRecentVals = sh.worksheet("Log - " + computerName).row_values(lastRow)
    mostRecentDate = datetime.datetime.strptime(mostRecentVals[0], r"%m/%d/%Y").date()
    rng = list(pd.date_range(start=mostRecentDate, end=datetime.datetime.today(), freq='D'))
    rng.pop(0)
    dDif = len(rng)
    rng = [ts.date().strftime("%m/%d/%Y") for ts in rng]
    update = []
    if len(rng) != 0:
        del rng[-1]
        dDif = len(rng)
        for i, _ in enumerate(rng):
            update.append([rng[i], 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0])
    else:
        lastRow -= 1
    nowLastRow = lastRow + 1 + dDif
    # print(f"nowLastRow: {nowLastRow}")
    nowLastRowVals = sh.worksheet("Log - " + computerName).row_values(nowLastRow)
    # print(f"nowLastRowVals: {nowLastRowVals}")
    nowLastRowVals = [0 if val == '' else val for val in nowLastRowVals]
    while len(nowLastRowVals) < 21:
        nowLastRowVals.append(0)
    while len(nowLastRowVals) > 21:
        newLen = len(nowLastRowVals)
        del nowLastRowVals[newLen - 1]
    # print(nowLastRowVals)
    update.append([
        datetime.datetime.now().strftime("%m/%d/%Y"),
        r['desktop-files-deleted'] + int(nowLastRowVals[1]),
        r['desktop-folders-deleted'] + int(nowLastRowVals[2]),
        r['document-files-deleted'] + int(nowLastRowVals[3]),
        r['document-folders-deleted'] + int(nowLastRowVals[4]),
        r['pictures-files-deleted'] + int(nowLastRowVals[5]),
        r['pictures-folders-deleted'] + int(nowLastRowVals[6]),
        r['videos-files-deleted'] + int(nowLastRowVals[7]),
        r['videos-folders-deleted'] + int(nowLastRowVals[8]),
        r['music-files-deleted'] + int(nowLastRowVals[9]),
        r['music-folders-deleted'] + int(nowLastRowVals[10]),
        r['downloads-files-deleted'] + int(nowLastRowVals[11]),
        r['downloads-folders-deleted'] + int(nowLastRowVals[12]),
        r['chrome-files-deleted'] + int(nowLastRowVals[13]),
        r['chrome-folders-deleted'] + int(nowLastRowVals[14]),
        r['chrome-instances-killed'] + int(nowLastRowVals[15]),
        r['edge-files-deleted'] + int(nowLastRowVals[16]),
        r['edge-folders-deleted'] + int(nowLastRowVals[17]),
        r['edge-instances-killed'] + int(nowLastRowVals[18]),
        r['programs-uninstalled'] + int(nowLastRowVals[19]),
        1 + int(nowLastRowVals[20])
    ])
    sh.worksheet("Log - " + computerName).update(f'A{(lastRow + 1)}:U{(lastRow + 1 + dDif)}', update)

def getMonthlyVals(api_key, from_email, to_email, subject):
    worksheets = sh.worksheets()
    # print(worksheets)
    computers = []
    for i, ws in enumerate(worksheets):
        if i + 1 < len(worksheets) and worksheets[i + 1].title.find("Log") != -1:
            computers.append(ws.title)
    # print(computers)
    res = []
    for comp in computers:
        vals = sh.worksheet(comp).get('H9:J23')
        res.append({
            'name': comp,
            'desktop-files-deleted': vals[0][0],
            'desktop-folders-deleted': vals[0][1],
            'document-files-deleted': vals[1][0],
            'document-folders-deleted': vals[1][1],
            'pictures-files-deleted': vals[2][0],
            'pictures-folders-deleted': vals[2][1],
            'videos-files-deleted': vals[3][0],
            'videos-folders-deleted': vals[3][1],
            'music-files-deleted': vals[4][0],
            'music-folders-deleted': vals[4][1],
            'downloads-files-deleted': vals[5][0],
            'downloads-folders-deleted': vals[5][1],
            'chrome-files-deleted': vals[6][0],
            'chrome-folders-deleted': vals[6][1],
            'chrome-instances-killed': vals[11][2],
            'edge-files-deleted': vals[7][0],
            'edge-folders-deleted': vals[7][1],
            'edge-instances-killed': vals[12][2],
            'programs-uninstalled': vals[13][2],
            'times-ran': vals[14][2],
        })
        clearMonthlyStats(comp)
    sendMonthlyReport(api_key, from_email, to_email, subject, res)

# r = results
def sendResults(computerName, r):
    a = sh.worksheet(computerName).get('H9:J23')
    sh.worksheet(computerName).update('H9:I16', [
        [(r['desktop-files-deleted'] + int(a[0][0])), (r['desktop-folders-deleted'] + int(a[0][1]))],
        [(r['document-files-deleted'] + int(a[1][0])), (r['document-folders-deleted'] + int(a[1][1]))],
        [(r['pictures-files-deleted'] + int(a[2][0])), (r['pictures-files-deleted'] + int(a[2][1]))],
        [(r['videos-files-deleted'] + int(a[3][0])), (r['videos-folders-deleted'] + int(a[3][1]))],
        [(r['music-files-deleted'] + int(a[4][0])), (r['music-folders-deleted'] + int(a[4][1]))],
        [(r['downloads-files-deleted'] + int(a[5][0])), (r['downloads-folders-deleted'] + int(a[5][1]))],
        [(r['chrome-files-deleted'] + int(a[6][0])), (r['chrome-folders-deleted'] + int(a[6][1]))],
        [(r['edge-files-deleted'] + int(a[7][0])), (r['edge-folders-deleted'] + int(a[7][1]))],
    ])
    sh.worksheet(computerName).update('J20:J23', [
        [(r['chrome-instances-killed'] + int(a[11][2]))],
        [(r['edge-instances-killed'] + int(a[12][2]))],
        [(r['programs-uninstalled'] + int(a[13][2]))],
        [(r['times-ran'] + int(a[14][2]))],
    ])
    sh.worksheet(computerName).update('B5', str(datetime.datetime.now()))