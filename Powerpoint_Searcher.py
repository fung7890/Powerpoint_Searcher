from pptx import Presentation
from os import listdir
import os
from tkinter import *
from tkinter import ttk
import tkinter as tk
from tkinter import filedialog
import tkinter.scrolledtext as tkst

directoryPath = ''
text_runs = []

def findWord(filePath):
    search = searchWord.get()
    prs = Presentation(filePath)

    for count, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    
                    if search.lower() in run.text.lower():
                        text_runs.append([run.text, count, os.path.basename(os.path.dirname(filePath)), os.path.basename(filePath), filePath])

def findFiles(path):
    #path = r'C:\Users\owner\Dropbox\Python\Projects\pptx_folder'
    all_files = []
    for dir, subdir, files in os.walk(path):

        for file in files:
            filename, file_ext = os.path.splitext(file)

            if file_ext == ".pptx":               
                all_files.append((os.path.join(dir, file)))
        
        if len(subdir) > 0 and subDirectory.get():
            del subdir[:]

    return all_files

def main():
    global textBox
    textBox.config(state=NORMAL)    
    text_runs.clear()
    for i in findFiles(directoryPath):
        findWord(i)
    output(text_runs)

    

root = tk.Tk()
root.title("Powerpoint Searcher")

showDir = StringVar()
showDir.set("Current Folder: "+directoryPath)

subDirectory = tk.IntVar()
subDirectory.set(0)

def open():
    global directoryPath
    global showDir
    directoryPath =  filedialog.askdirectory()
    directoryPath = (os.path.abspath(directoryPath))
    showDir.set("Current Folder: "+directoryPath)

def output(text_runs):
    global textBox
    textBox.delete('1.0', END)
    for count, found in enumerate(text_runs):
        textBox.insert('1.0', "SLIDE: " + str(found[1]) + " "*5 + " FILE: " + found[3] + " "*5 + " FOLDER: " + found[2], count)
        textBox.insert('1.0', '\n')
        textBox.tag_config(count, foreground="blue", font="calibri")
        textBox.tag_bind(count, "<ButtonRelease-1>", lambda event, i=found: startApp(Event, i))
    textBox.config(state=DISABLED)

def startApp(event, found):
    os.startfile(found[4])



mainframe = ttk.Frame(root, padding="23 23 42 42")
mainframe.grid(column=0, row=0)

searchWord= ttk.Entry(mainframe, width=100)
searchWord.grid(column=1, row=5, sticky=(W, E))

ttk.Button(mainframe, text="Search", command = main, width=20).grid(column=2, row=5)
ttk.Button(mainframe, text="Open Folder", command = open, width=20).grid(column=2, row=8)

ttk.Checkbutton(mainframe, text="Do not include folders in folder", onvalue=1, offvalue=0, variable=subDirectory).grid(column=2, row=9)

ttk.Label(mainframe, textvariable=showDir).grid(column=1, row=7, sticky=(W, E))

textBox = tkst.ScrolledText(mainframe, cursor="arrow", width = 100)
textBox.grid(column=1, row=10, columnspan=2)
textBox.config(state=DISABLED)    

root.mainloop()






















'''
TODO:
access within directories of directories

corner cases for search:
    ie. spelled it wrong, embeded in a chart?

UI: Tkinter has open directory, check other UI libraries 

Testing:
    500 files at once 


'''
