import win32api
import win32print
import traceback
import time
import os
import shutil
from os import path
from tkinter.filedialog import askdirectory
from tkinter import *
from tkinter import font # * doesn't import font or messagebox
from tkinter import messagebox
from glob import glob
from subprocess import call

_printer = "Canon MF240 Series UFRII LT"
_picklistIncomingFolder = "Batch In"
_picklistOutgoingFolder = "Batch Out"
_picklistIncomingPDFs = glob(path.join(_picklistIncomingFolder,"*.{}".format("pdf")))
_picklists = [s for s in _picklistIncomingPDFs if "picklist" in s.lower()]

_labelers = ["ZDesigner GK420d","ZDesigner GX420d"]
_labelIncomingFolder = "Batch In"
_labelOutgoingFolder = "Batch Out"
_labelIncomingPDFs = glob(path.join(_labelIncomingFolder,"*.{}".format("pdf")))
_labels = [s for s in _labelIncomingPDFs if "labels" in s.lower()]
_labelfilecount = len(_labelfiles)

GHOSTSCRIPT_PATH = 'GHOSTSCRIPT\\bin\\gswin32.exe'
GSPRINT_PATH = 'GSPRINT\\gsprint.exe'

def printBatch(labeler, labels, printer, picklist):

    print("Printing picklist...")
    win32api.ShellExecute(0, 'open', GSPRINT_PATH, '-ghostscript "'+GHOSTSCRIPT_PATH+'" -printer "'+printer+'" "'+picklist+'"', '.', 0)
    time.sleep(1)

    print("Printing labels...")
    win32api.ShellExecute(0, 'open', GSPRINT_PATH, '-ghostscript "'+GHOSTSCRIPT_PATH+'" -printer "'+labeler+'" "'+labels+'"', '.', 0)
    time.sleep(1)
    print("Done!")

#==============================#
#=============MAIN=============#
#==============================#
# remember: run with F5

print("Printer:",_printer)
print('Picklist Folder:', _picklistIncomingFolder)
print('Picklist Count:', str(len(_picklists)))
print("")

print("Labelers:",_labelers)
print('Label Folder:', _labelIncomingFolder)
print('Label File Count:', str(len(_labels)))
print("")

res = input("Press enter to start.")
if res == "":
    if(len(_picklists) != len(_labelfiles)):
        print("Error, picklists and batches are not equal.")
        exit(-69)

    x = 0
    while x < len(_picklists):
        picklistsToMove = []
        labelsToMove = []
        for _labeler in _labelers:
            print("")
            print("Printing file pair: ",str(x + 1))

            picklist = [s for s in _picklists if "batch " + str(x + 1) + " labels" in s.lower()][0]
            if not picklist:
                res = input("Picklist for batch " + str(x) + " not found. Continue to next batch? Y/N")
                if res == "Y":
                    continue
                else:
                    exit(420)

            labels = [s for s in _labels if "batch " + str(x + 1) + " labels" in s.lower()][0]
            if not labels:
                res = input("Labels for batch " + str(x) + " not found. Continue to next batch? Y/N")
                if res == "Y":
                    continue
                else:
                    exit(420)

            print("Printer:",_printer)
            print("Picklist:",picklist)
            print("Labeler:",_labeler)
            print("Labels:",labels)

            printBatch(_labeler, labels, _printer, picklist)

            picklistsToMove.append(picklist)
            labelsToMove.append(labels)

            x += 1
            if x >= len(_picklists):
                break

        res = input("Did the files print successfully? Y/N")
        if res == "Y":
            # move files to printed folder
            print("Moving files to printed folder...")
            for p in picklistsToMove:
                shutil.move(p, _picklistOutgoingFolder + "/" + os.path.basename(p))
            for l in labelsToMove:
                shutil.move(l, _labelOutgoingFolder + "/" + os.path.basename(l))
            print("Done!")
        else:
            print("Well shit, that's too bad.")
            exit(-69)
        if x < len(_picklists) - 1:
            res = input("Continue to next files? Y/N")
            if res == "Y":
                continue
            else:
                exit(420)

    print("All files printed.")
    exit(420)
else:
    print("Pussy Bitch.")
    exit(420)
