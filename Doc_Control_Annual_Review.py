import pyautogui ,time, os
from tkinter import *
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd


def windowsizecheck():
    x = pyautogui.size().width
    y = pyautogui.size().height
    print(x, y)
    if x == 1920 and y == 1080:
        print('All Good')
    else:
        messagebox.showinfo('Window Size Error', 'Ensure your screen resolution is set to 1920x1080')
        close_window

windowsizecheck()
root = Tk()
root.geometry("410x200")
root.title('Doc Review SignOff')
myfile = StringVar(root)

def choose_myfile(*_):
    global myfile
    myfile = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("excel files","*.xlsx"),("all files","*.*")))
    print(myfile)
    RunButton.config(state='normal')
    return myfile

def moveto(xPos, yPos):
    pyautogui.moveTo(xPos, yPos, duration=0.25)

def click(clickcount):
    pyautogui.click(clicks=clickcount)

def entertext(text):
    pyautogui.typewrite(text)
    pyautogui.typewrite(['enter'])

def deletetext():
    pyautogui.typewrite(['backspace', 'backspace', 'backspace', 'backspace', 'backspace', 'backspace','backspace', 'backspace', 'backspace', 'backspace', 'backspace', 'backspace'])

def DocReviewSO(doc, date):
    time.sleep(.5)

    moveto(433, 206) # Move to Record Text Box
    click(1) # Place cursor at text box
    deletetext() # Ensure no txt remains in box
    entertext(doc) # Enter Ticket ID number
    time.sleep(1) 
    moveto(286,359) # Move to first record
    click(2) # Open first record
    moveto(548,220) # Move to Reviews Section
    click(1) # Click Tab
    moveto(1860,247) # Move to Signoff Section
    click(1)
    time.sleep(0.5)
    moveto(992,525)
    entertext(date) # Enter Completed Date
    pyautogui.hotkey('ctrl','s') # Control + S to save
    moveto(279,139) # Search for next recrod
    click(1)

# Close the program window
def close_window():
    root.destroy()


#Program Loop Through Files
def uniLoop():
    df = pd.read_csv(myfile)
    count = 0
    #DocReviewSO('ANO-W-1592', '5/20/2019')  # UnComment this section if running a single record
    for index, row in df.iterrows(): # Use loop if running multiple records.
        print (row['Doc_Num'], row['Approval_Date'])
        count += 1
        DocReviewSO(row['Doc_Num'], row['Approval_Date'])

    messagebox.showinfo('Message', '{} records updated'.format(count))



# Program Window with Widgets
RunButton = ttk.Button(root, text='Run', command=uniLoop, state=DISABLED)
FilButton = ttk.Button(root, text='Select File', command=choose_myfile)
StartText = ttk.Label(root, text='1. Ensure screen resolution is set to 1920x1080.')
StartText2 = ttk.Label(root, text='2. Have uniPoint Documents window open and maximized.')
StartText3 = ttk.Label(root, text="3. After clicking OK you'll have 5 secs to bring the uniPoint screen into focus.")


StartText.grid(column=0, row=0, columnspan=3, sticky=W)
StartText2.grid(column=0, row=1, columnspan=3, sticky=W)
StartText3.grid(column=0, row=2, columnspan=3, sticky=W)
RunButton.grid(column=1, row=4, sticky=E)
FilButton.grid(column=1, row=4, sticky=W)
root.mainloop()