import win32api
import win32print
import traceback

from tkinter.filedialog import askopenfilename
from tkinter import *
from tkinter import font, filedialog  # * doesn't import font or messagebox
from tkinter import messagebox

root = Tk()
root.title("Python Printer")
root.geometry("410x310")
root.resizable(False, False)
root.tk.call('encoding', 'system', 'utf-8')


def font_size(fs):
    return font.Font(family='Helvetica', size=fs, weight='bold')


# Add a grid
mainframe = Frame(root)
# mainframe.grid(column=0,row=0, sticky=(N,W,E,S) )
mainframe.grid(column=0, row=0, sticky=(N))
mainframe.columnconfigure(0, weight=1)
mainframe.rowconfigure(0, weight=1)
mainframe.pack(pady=10, padx=0)

# Create a _printer variable
_picklistPrinter = StringVar(root)
_labelPrinter = StringVar(root)
_filename = ""


# on change dropdown value
def sel_picklist_printer(*args):
    print(_picklistPrinter.get())


# on change dropdown value
def sel_label_printer(*args):
    print(_picklistPrinter.get())


# link function to change dropdown
_picklistPrinter.trace('w', sel_picklist_printer)
_labelPrinter.trace('w', sel_label_printer)


def upload(event=None):
    global _filename
    _filename = filedialog.askopenfilename()
    # print('Selected:', _filename)


def print_action(event=None):

    if not _filename:
        messagebox.showerror("Error", "No File Selected")
        return
    elif not _picklistPrinter.get():
        messagebox.showerror("Error", "No Printer Selected")
        return

    try:
        # win32print.SetDefaultPrinter(_printer.get())
        win32api.ShellExecute(0, "print", _filename, None, ".", 0)
        win32print.ClosePrinter(pHandle)
    except:
        pass
        messagebox.showerror("Error", "There was an error printing the file :(")


choices = [printer[2] for printer in win32print.EnumPrinters(2)]
popupMenu = OptionMenu(mainframe, _picklistPrinter, *choices)
popupMenu['font'] = font_size(12)
Label(mainframe, text="SELECT PICKLIST PRINTER").grid(row=1, column=1)
popupMenu.grid(row=2, column=1)

choices = [printer[2] for printer in win32print.EnumPrinters(2)]
popupMenu2 = OptionMenu(mainframe, _picklistPrinter, *choices)
popupMenu2['font'] = font_size(12)
Label(mainframe, text="SELECT LABEL PRINTER").grid(row=3, column=1)
popupMenu2.grid(row=4, column=1)

Label(mainframe, text="SELECT INCOMING FOLDER").grid(row=5, column=1)
button = Button(mainframe, text=u"\uD83D\uDCC1" ' BROWSE', command=upload)
button['font'] = font_size(12)
button.grid(row=6, column=1)

Label(mainframe, text="SELECT OUTGOING FOLDER").grid(row=7, column=1)
button = Button(mainframe, text=u"\uD83D\uDCC1" ' BROWSE', command=upload)
button['font'] = font_size(12)
button.grid(row=8, column=1)

Label(mainframe).grid(row=10, column=1)
p_button = Button(mainframe, text=u'\uD83D\uDDB6' + " PRINT", command=print_action, fg="dark green", bg="white")
p_button['font'] = font_size(18)
p_button.grid(row=11, column=1)

root.mainloop()
