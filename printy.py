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
_picklistIncomingFolder = "Picklists In"
_picklistOutgoingFolder = "Picklists Out"
_printerfiles = glob(path.join(_picklistIncomingFolder,"*.{}".format("pdf")))
_printerfilecount = len(_printerfiles)

_labelers = ["ZDesigner GK420d","ZDesigner GX420d"]
_labelIncomingFolder = "Labels In"
_labelOutgoingFolder = "Labels Out"
_labelfiles = glob(path.join(_labelIncomingFolder,"*.{}".format("pdf")))
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
print('Printer Folder:', _picklistIncomingFolder)
print('Printer File Count:', str(_printerfilecount))
print("")

print("Labelers:",_labelers)
print('Label Folder:', _labelIncomingFolder)
print('Label File Count:', str(_labelfilecount))
print("")

res = input("Press enter to start.")
if res == "":
    if(len(_printerfiles) != len(_labelfiles)):
        print("Error, picklists and batches are not equal.")
        exit(-69)

    x = 0
    while x < len(_printerfiles):
        picklistsToMove = []
        labelsToMove = []
        for _labeler in _labelers:
            print("")
            print("Printing file pair: ",str(x + 1))

            picklist = [s for s in _printerfiles if "Batch " + str(x + 1) + " Picklist" in s][0]
            if not picklist:
                print("Error during fetching picklist.")
                exit(-69)

            labels = [s for s in _labelfiles if "Batch " + str(x + 1) + " labels" in s][0]
            if not labels:
                print("Error during fetching labels.")
                exit(-69)

            print("Printer:",_printer)
            print("Picklist:",picklist)
            print("Labeler:",_labeler)
            print("Labels:",labels)

            printBatch(_labeler, labels, _printer, picklist)

            picklistsToMove.append(picklist)
            labelsToMove.append(labels)

            x += 1
            if x >= len(_printerfiles):
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
        if x < len(_printerfiles) - 1:
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
