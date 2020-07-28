import re
import subprocess
from colorama import Fore, Style
import logging

#############
# PROGRAM INFORMATION
#############
VERSION = "1.0.0"
AUTHOR = "cbrxyz"
REPO = "https://github.com/cbrxyz/pcc"
#############

logging.basicConfig(
    filename="pcc.log",
    filemode='a',
    format='%(asctime)s.%(msecs)d %(levelname)s %(message)s',
    datefmt='%m-%d-%y %H:%M:%S',
    level=logging.DEBUG
)

def getVariable(variable):
    f = open("settings.txt")
    for l in f:
        matches = re.findall(r'(?<=' + variable + r'=).*', l)
        if len(matches) > 0: return matches[0]
    f.close()
    return False

def execute(cmd, regex, straightOutput=False):
    output = subprocess.Popen(cmd, stdout=subprocess.PIPE, shell=True).communicate()[0]
    output = output.decode('UTF-8')
    if straightOutput:
        return output
    else:
        numDeleted = len(re.findall(regex, output))
        return numDeleted

def log(text):
    logging.debug(text)
    print(text)

def warn(text):
    logging.warning(text)
    print(Fore.YELLOW + " [WARN] " + text + Style.RESET_ALL)

def success(text):
    logging.info(text)
    print(Fore.GREEN + Style.BRIGHT + text + Style.RESET_ALL)

def accentuate(header):
    logging.info(f"===== {header} =====")
    print(Fore.CYAN + Style.BRIGHT + "=====\n" + header + "\n=====" + Style.RESET_ALL)