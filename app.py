from tkinter import*
from tkinter import ttk
from tkinter import messagebox,filedialog
import pyautogui as at, pandas as pd
import threading,csv

from tkinter import filedialog
normal = ('Angsana New',15)
font1 = ('Angsana New',25)
font2 = ('Angsana New',20)
programname = "Bot 1688"

def getSettingsProgram():
    datas = open("./dataProgram/setting.txt",'r',encoding="utf8")
    result = {}
    for data in datas.readlines():
        splitdata = data.split('=')
        if(len(splitdata)==2):
            result[splitdata[0]] = splitdata[1].split(',')[0]
    return result
def setSettingsProgram(pathFile,newSetting):
    # ./dataProgram/setting.txt
    datas = open(pathFile,'w',encoding="utf8")
    datas.write(newSetting)
    datas.close()
def selectPath():
    folder_selected = filedialog.askdirectory()
    setSettingsProgram(folder_selected)

# textPath = StringVar()
# textPath.set(getSettingsProgram()["pathdownload"])
fside10 = ("Angsana New",10)
# font1 = ("Angsana New",10)
app = Tk()
app.title(programname)
app.config(bg="#ffffff") 
app.geometry("500x800-0+0")

titlePath = Label(app,text=getSettingsProgram()["pathdownload"])
titlePath.place(x=50,y=20)
btnselectPath = Button(app,text="Path",command=selectPath)
btnselectPath.place(x=10,y=20)

app.mainloop()
