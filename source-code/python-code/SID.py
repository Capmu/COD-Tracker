'''
    project   : SID : Same Information Detector
    version   : 2020.1
    developer : Capmu

    Unique Statement
        . . .
'''
#--------------------------------------------------------
# Classes / Libraries / Packages
#--------------------------------------------------------
from Classes.DeliveryInfo import DeliveryInfo

import os
import xlrd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from shutil import copyfile, move
import datetime
import time

#--------------------------------------------------------
# Fucntions
#--------------------------------------------------------
def getAllFileNameAtPath(path):
    for filePath in os.walk(path): #full example --> for root, dirs, files in os.walk("./Checking File"):
        for files in filePath:
            pass

    return(files)

#--------------------------------------------------------
# Variables
#--------------------------------------------------------
path = "source-code/python-code/SID-files"

#--------------------------------------------------------
# Implementation
#--------------------------------------------------------
fiels = getAllFileNameAtPath(path)
