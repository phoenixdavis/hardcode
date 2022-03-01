import json

import win32api
import win32print
import shutil

from tkinter import *
from tkinter import font, filedialog  # * doesn't import font or messagebox
from tkinter import messagebox
import os
from os import path
from glob import glob
import time

root = Tk()
root.title("Hard Printer")
root.geometry("610x510")
root.resizable(False, False)
root.tk.call('encoding', 'system', 'utf-8')

# Add a grid
mainframe = Frame(root)
# mainframe.grid(column=0,row=0, sticky=(N,W,E,S) )
mainframe.grid(column=0, row=0, sticky=N)
mainframe.columnconfigure(0, weight=1)
mainframe.rowconfigure(0, weight=1)
mainframe.pack(pady=10, padx=0)

# Create global variables
_incomingFolder = ""
_outgoingFolder = ""
_picklistPrinters = []
_labelPrinters = []
_picklists = []
_labels = []

def font_size(fs):
    return font.Font(family='Helvetica', size=fs, weight='bold')


# Selects the pprinters to use
def sel_picklist_printers(*args):
    global _picklistPrinters
    _picklistPrinters = [pPrinterListBox.get(i) for i in pPrinterListBox.curselection()]


# Selects the lprinters to use
def sel_label_printers(*args):
    global _labelPrinters
    _labelPrinters = [lPrinterListBox.get(i) for i in lPrinterListBox.curselection()]


# Selects folder containing batches to print
def sel_incoming_folder(event=None):
    global _incomingFolder, _picklists, _labels
    _incomingFolder = filedialog.askdirectory()
    incoming_pdfs = glob(path.join(_incomingFolder, "*.{}".format("pdf")))
    _picklists = [s for s in incoming_pdfs if "picklist" in s.lower()]
    _labels = [s for s in incoming_pdfs if "labels" in s.lower()]


# Selects folder to send printed batches
def sel_outgoing_folder(event=None):
    global _outgoingFolder
    _outgoingFolder = filedialog.askdirectory()


def print_action(event=None):
    if not _incomingFolder or not _outgoingFolder:
        messagebox.showerror("Error", "Folders Not Selected")
        return
    elif len(_picklistPrinters) == 0:
        messagebox.showerror("Error", "No picklist printer selected.")
        return
    elif len(_labelPrinters) == 0:
        messagebox.showerror("Error", "No label printer selected.")
        return
    elif len(_picklistPrinters) < len(_labelPrinters) / 2:
        messagebox.showerror("Error", "Must have at least one picklist printer for every two label printers.")
        return
    elif len(_picklists) != len(_labels):
        messagebox.showerror("Error", "Picklists/Labels are not equal.")
        return
    try:
        message = 'Picklist Printers: ' + str(_picklistPrinters) + \
                  '\n\n' + 'Label Printers: ' + str(_labelPrinters) + \
                  '\n\n' + 'Batch Count: ' + str(len(_picklists)) + \
                  '\n\n' + 'Incoming Folder: ' + str(_incomingFolder) + \
                  '\n\n' + 'Outgoing Folder: ' + str(_outgoingFolder) + \
                  '\n\nBegin printing process? '
        answer = messagebox.askyesno(title='Confirm', message=message)
        if answer:
            print_batches()
    except:
        pass
        messagebox.showerror("Error", "There was an error printing the file :(")


def print_batch(labeler, label, printer, picklist):
    print("Printing file pair: ", picklist + ' ' + label)
    GHOSTSCRIPT_PATH = 'GHOSTSCRIPT\\bin\\gswin32.exe'
    GSPRINT_PATH = 'GSPRINT\\gsprint.exe'

    win32api.ShellExecute(
        0,
        'open',
        GSPRINT_PATH,
        '-ghostscript "' + GHOSTSCRIPT_PATH + '" -printer "' + printer + '" "' + picklist + '"',
        '.',
        0
    )
    time.sleep(1)

    win32api.ShellExecute(
        0,
        'open',
        GSPRINT_PATH,
        '-ghostscript "' + GHOSTSCRIPT_PATH + '" -printer "' + labeler + '" "' + label + '"',
        '.', 0
    )
    time.sleep(1)


def print_batches():
    picklists = _picklists
    labels = _labels

    # while we have shit to print
    while len(picklists) > 0:
        availPPrints = _picklistPrinters.copy()
        availLPrints = _labelPrinters.copy()
        picklists_to_move = []
        labels_to_move = []

        # while we have 1 pprint and 2 lprints
        while len(availPPrints) > 0 and len(availLPrints) > 0:

            # fetch printers
            pp = availPPrints.pop()
            lp1 = availLPrints.pop()

            # fetch next batch
            picklist = picklists.pop(0)
            label = labels.pop(0)

            # execute print
            print_batch(lp1, label, pp, picklist)

            picklists_to_move.append(picklist)
            labels_to_move.append(label)

            # if we have another batch and lprinter, use the same pprinter before continuing
            if len(availLPrints) > 0 and len(picklists) > 0:
                lp2 = availLPrints.pop()
                picklist = picklists.pop(0)
                label = labels.pop(0)

                print_batch(lp2, label, pp, picklist)

                picklists_to_move.append(picklist)
                labels_to_move.append(label)

        answer = messagebox.askyesno(title='Confirm', message='Did the files print successfully?')
        if answer:
            # move files to printed folder
            for pick in picklists_to_move:
                shutil.move(pick, str(_outgoingFolder) + "/" + os.path.basename(pick))
            for lab in labels_to_move:
                shutil.move(lab, str(_outgoingFolder) + "/" + os.path.basename(lab))
            messagebox.showinfo(title='Info', message='Files moved to printed folder.')
        else:
            return

        if len(picklists) > 0:
            answer = messagebox.askyesno(title='Confirm', message='Continue to next cycle?')
            if not answer:
                return

    messagebox.showinfo(title='Complete', message='All files printed.')


def load_settings():
    global _incomingFolder, _outgoingFolder, _picklistPrinters, _labelPrinters, _picklists, _labels
    file = open("config.json", "r")
    if file:
        settings = json.load(file)
        if settings:
            _incomingFolder = settings["INCOMINGFOLDER"]
            _outgoingFolder = settings["OUTGOINGFOLDER"]
            _picklistPrinters = settings["PICKLISTPRINTERS"]
            _labelPrinters = settings["LABELPRINTERS"]
            incoming_pdfs = glob(path.join(_incomingFolder, "*.{}".format("pdf")))
            _picklists = [s for s in incoming_pdfs if "picklist" in s.lower()]
            _labels = [s for s in incoming_pdfs if "labels" in s.lower()]

            message = 'Settings Loaded: ' + \
                      '\n\nPicklist Printers: ' + str(_picklistPrinters) + \
                      '\n\nLabel Printers: ' + str(_labelPrinters) + \
                      '\n\nIncoming Folder: ' + str(_incomingFolder) + \
                      '\n\nOutgoing Folder: ' + str(_outgoingFolder)
            messagebox.showinfo(title='Settings Loaded', message=message)
        else:
            messagebox.showerror("Error", "Error loading settings.")
    else:
        messagebox.showerror("Error", "Settings file not found.")


def save_settings():
    global _incomingFolder, _outgoingFolder, _picklistPrinters, _labelPrinters, _picklists, _labels
    settings = {
        "INCOMINGFOLDER": _incomingFolder,
        "OUTGOINGFOLDER": _outgoingFolder,
        "PICKLISTPRINTERS": _picklistPrinters,
        "LABELPRINTERS": _labelPrinters
    }
    file = open("config.json", "w")

    json.dump(settings, file)

    file.close()
    message = 'Settings Saved: ' + \
              '\n\nPicklist Printers: ' + str(_picklistPrinters) + \
              '\n\nLabel Printers: ' + str(_labelPrinters) + \
              '\n\nIncoming Folder: ' + str(_incomingFolder) + \
              '\n\nOutgoing Folder: ' + str(_outgoingFolder)
    messagebox.showinfo(title='Settings Saved', message=message)


Label(mainframe, text="SELECT PICKLIST PRINTERS").grid(row=1, column=1)
choices = [printer[2] for printer in win32print.EnumPrinters(2)]
pPrinterListBox = Listbox(mainframe, selectmode="multiple", height=5, width=40)
pPrinterListBox.grid(row=2, column=1, columnspan=1)
for each_item in range(len(choices)):
    pPrinterListBox.insert(END, choices[each_item])

button = Button(mainframe, text='Set Picklist Printers', command=sel_picklist_printers)
button['font'] = font_size(12)
button.grid(row=2, column=2)


Label(mainframe, text="SELECT LABEL PRINTERS").grid(row=3, column=1, pady=(20, 0))
lPrinterListBox = Listbox(mainframe, selectmode="multiple", height=5, width=40)
lPrinterListBox.grid(row=4, column=1, columnspan=1)
for each_item in range(len(choices)):
    lPrinterListBox.insert(END, choices[each_item])

button = Button(mainframe, text='Set Label Printers', command=sel_label_printers)
button['font'] = font_size(12)
button.grid(row=4, column=2)


Label(mainframe, text="SELECT INCOMING FOLDER").grid(row=5, column=1, pady=(20, 0))
button = Button(mainframe, text=u"\uD83D\uDCC1" ' BROWSE', command=sel_incoming_folder)
button['font'] = font_size(12)
button.grid(row=6, column=1)


Label(mainframe, text="SELECT OUTGOING FOLDER").grid(row=5, column=2, pady=(20, 0))
button = Button(mainframe, text=u"\uD83D\uDCC1" ' BROWSE', command=sel_outgoing_folder)
button['font'] = font_size(12)
button.grid(row=6, column=2)


p_button = Button(mainframe, text="LOAD", command=load_settings, fg="dark green", bg="white")
p_button['font'] = font_size(18)
p_button.grid(row=7, column=1, pady=(20, 0))

p_button = Button(mainframe, text="SAVE", command=save_settings, fg="dark green", bg="white")
p_button['font'] = font_size(18)
p_button.grid(row=7, column=2, pady=(20, 0))


Label(mainframe).grid(row=8, column=1, pady=(20, 0))
p_button = Button(mainframe, text=u'\uD83D\uDDB6' + " PRINT", command=print_action, fg="dark green", bg="white")
p_button['font'] = font_size(18)
p_button.grid(row=9, column=1, columnspan=2)

root.mainloop()
