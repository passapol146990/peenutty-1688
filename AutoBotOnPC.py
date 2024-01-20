from tkinter import*
from tkinter import ttk
from tkinter import messagebox,filedialog
import pyautogui as at, pandas as pd
import threading,csv

# webdriver
from selenium import webdriver
from selenium.webdriver.common.by import By
import bs4,time,pyperclip
from selenium import webdriver



############## อ่านจาก NameUrlSave.txt เพื่อมาตั้งชื่อไฟล์ .txt เก็บ url ######################
def ReadNameUrlSave():
    NameUrlSaveValues = []
    try:
        NameUrlSave = open('NameUrlSave.txt','r',encoding='utf-8').readlines()
        for i in NameUrlSave:
            NameUrlSaveValues.append(i.split('\n')[0])
    except:
        open('NameUrlSave.txt','w',encoding='utf-8')
    return NameUrlSaveValues
############## ตัวแปรชุดข้อมูล ######################
normal = ('Angsana New',15)
font1 = ('Angsana New',25)
font2 = ('Angsana New',20)
programname = "Auto Bot on PC v.1.9"
############## windown ######################
window = Tk()
# 400x400 ขนาดหน้าจอ +500+150 ตำแน่งโปรแกรม x, y
window.geometry('720x680+680+150') 
window.title(programname)
############### Notbook #####################
# สร้างแทบด้านบน t1 -t3 เป็น gui ได้เลย
tab = ttk.Notebook(window)
t1 = Frame(tab)
t2 = Frame(tab)
t3 = Frame(tab)

tab.pack(fill=BOTH,expand=1)

tab.add(t1,text='home')
tab.add(t2,text='View System commands')
tab.add(t3,text='Website')


##############  หน้า1 ################
divPage1 = Frame(t1)
divPage2 = Frame(t1)
divPage1.pack(side=LEFT,fill=BOTH)
divPage2.pack(side=LEFT,fill=BOTH,expand=True)
div2_1 = Frame(divPage2)
div2_2 = Frame(divPage2)
div2_3 = Frame(divPage2)
div2_1.pack(side=TOP,fill=BOTH,expand=True)
div2_3.pack(side=BOTTOM,fill=BOTH,expand=True)
    
def position():
    for i in range(3,0,-1):
        ganXY.set(f'{i}')
        time.sleep(1)
    x,y = at.position()
    ganXY.set(f'x = {x},y = {y}')
    
    option = title_option.get()
    if(option == "move mouse"):
        input_number.set(f"{x},{y}")

def positionMouse():
    t = threading.Thread(target=position)
    t.start()

nameFile = 'log.csv'
def save_command(selected_command):
    with open(nameFile, mode='a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow([selected_command])
    print(f"บันทึกคำสั่ง: {selected_command}")

# /////////////////////////////
def AddCommands():
    option = title_option.get()
    if(option == "time delay"):
        try:
            count = int(input_number.get())
            save_command(f'timesleep:{count}')
            messagebox.showinfo("success",f"บันทึก Time Delay จำนวน {count} วิ สำเร็จ")
        except:
            messagebox.showinfo("error","กรุณาใส่ตัวเลขเท่านั้น")
    elif(option == "left click"):
        try:
            count = int(input_number.get())
            save_command(f'leftClick:{count}')
            messagebox.showinfo("success",f"บันทึก Left Click จำนวน {count}คลิ๊กสำเร็จ")
        except:
            messagebox.showinfo("error","กรุณาใส่ตัวเลขไม่เกิน 5")
    elif(option == "right click"):
        save_command(f'rightClick:1')
        messagebox.showinfo("success",f"บันทึก Right Click สำเร็จ")
    elif(option == "move mouse"):
        text = input_number.get().split(',')
        if len(text)!=2:
            messagebox.showinfo("error","กรุณาใส่ตัวเลข 2 ตัวมี , ขั่น")
            return
        else:
            try:
                position = [int(i) for i in text]
                save_command(f'moveTo:{position[0]},{position[1]}')
                messagebox.showinfo("success",f"บันทึก Move Mouse x={position[0]}, x={position[1]} สำเร็จ")
            except:
                messagebox.showinfo("error","กรุณาใส่ตัวเลข 2 ตัวมี , ขั่น")
                return
    elif(option == "scroll mouse"):
        try:
            text = int(input_number.get())
            save_command(f'scroll:{text}')
            messagebox.showinfo("success",f"บันทึก Scroll Mouse {text}สำเร็จ")
        except:
            messagebox.showinfo("error","กรุณาใส่ตัวเลข 1 ตัวเป็นจำนวน (+)(-)")
            return
    elif(option == "enter"):
        save_command(f'enter:1')
        messagebox.showinfo("success",f"บันทึก Enter สำเร็จ")
    elif(option == "CTRL + key"):
        key = input_number.get()
        save_command(f'CTRL:{key}')
        messagebox.showinfo("success",f"บันทึก CTRL + {key} สำเร็จ")
    elif(option == "click20Item"):
        save_command(f'click20Item:0')
        messagebox.showinfo("success",f"บันทึก click20Item สำเร็จ")
    elif(option == "check 0"):
        save_command(f'check:0')
        messagebox.showinfo("success",f"บันทึก check 0 สำเร็จ")
    elif(option == "save url"):
        save_command(f'save url:0')
        messagebox.showinfo("success",f"บันทึก save url สำเร็จ")
    elif(option == "DownloadExcel"):
        save_command(f'DownloadExcel:0')
        messagebox.showinfo("success",f"บันทึก DownloadExcel สำเร็จ")
    elif(option == "ID"):
        save_command(f'ID:0')
        messagebox.showinfo("success",f"บันทึก ID สำเร็จ")
    setTreeCommand()

def setTreeCommand():
    try:
        pathFile.set(nameFile)
        table.delete(*table.get_children())
        with open(nameFile, newline='') as csv_file:
            csv_reader = csv.reader(csv_file)
            i = 0
            for row in csv_reader:
                if i!=0:
                    try:
                        if ':' in row[0]:
                            s1 = row[0].split(':')
                            if ',' in s1[1]:
                                numberState = s1[1].split(',')
                                table.insert('','end',values=(i,s1[0], f'x={numberState[0]},y={numberState[1]}'))
                            else:
                                table.insert('','end',values=(i,s1[0],s1[1]))
                    except:
                        pass
                i+=1
    except:
        with open(nameFile, mode='a', newline='') as file:
            writer = csv.writer(file)
            writer.writerow([programname])

def AddCommands_system():
    global statusRunCommand
    statusRunCommand = False
    t = threading.Thread(target=AddCommands)
    t.start()
# ///////////////////////////
def TestRunCommand():
    option = title_option.get()
    def count_down3():
        for i in range(3,0,-1):
            if statusRunCommand:
                testSys.set(f'ทดสอบการทำงาน ดีเลย์ {3} วิ')
                return
            testSys.set(f'ทดสอบการทำงาน ดีเลย์ {i} วิ')
            time.sleep(1)
        testSys.set(f'ทดสอบการทำงาน ดีเลย์ {3} วิ')
    
    if(option == "left click"):
        try:
            count = int(input_number.get())
            count_down3()
            for i in range(count):
                at.click()
        except:
            messagebox.showinfo("error","กรุณาใส่ตัวเลขไม่เกิน 5")
    elif(option == "right click"):
        count_down3()
        at.rightClick()
    elif(option == "move mouse"):
        count_down3()
        text = input_number.get().split(',')
        if len(text)!=2:
            messagebox.showinfo("error","กรุณาใส่ตัวเลข 2 ตัวมี , ขั่น")
            return
        else:
            try:
                position = [int(i) for i in text]
                at.moveTo(position[0],position[1])
            except:
                messagebox.showinfo("error","กรุณาใส่ตัวเลข 2 ตัวมี , ขั่น")
                return
    elif(option == "scroll mouse"):
        count_down3()
        try:
            text = int(input_number.get())
            at.scroll(text)
        except:
            messagebox.showinfo("error","กรุณาใส่ตัวเลข 1 ตัวเป็นจำนวน (+)(-)")
            return
    elif(option == "enter"):
        count_down3()
        at.press('enter')
    elif(option == "CTRL + key"):
        text = str(input_number.get())
        count_down3()
        try:
            at.hotkey('ctrl',text)
        except:
            pass
    elif(option == "click20Item"):
        req = web.select_item20()
    elif(option == "check 0"):
        req = web.checkZero()
    elif(option == "save url"):
        text = input_number.get()
        if len(text)==0:
            messagebox.showinfo("error","กรุณาใส่ชื่อไฟล์ที่ต้องการทดสอบ")
            return
        web.saveUrl(name=text)
    elif(option == "DownloadExcel"):
        req = web.DownloadExcel()
    elif(option == "ID"):
        text = int(input_number.get())
        if len(str(text))==0:
            messagebox.showinfo("error","กรุณาใส่ID")
            return
        try:
            print(ReadNameUrlSave()[text])
            pyperclip.copy(ReadNameUrlSave()[text])
        except:
            pyperclip.copy('indexError')

def Runtest_system():
    global statusRunCommand
    statusRunCommand = False
    t = threading.Thread(target=TestRunCommand)
    t.start()

def stoptest_system():
    global statusRunCommand
    statusRunCommand = True
    messagebox.showinfo("หยุดทำงาน test",f"หยุดทำงาน test")

loop = StringVar()
ReadStepRun = StringVar()
pathFile = StringVar()
idStep = StringVar()
title_option = StringVar()
title_input_number = StringVar()
input_number = StringVar()
ganXY = StringVar()
testSys = StringVar()
testSys.set(f'ทดสอบการทำงาน ดีเลย์ {3} วิ')
title_option.set("<--กรุณาเลือกคำสั่งจากด้านซ้าย")

def setOption(title):
    form1.pack()
    input_number.set("")
    if len(title_option.get())>0:
        addCommand.pack(expand=True,ipadx=50,ipady=20,padx=20)
    title_option.set(title)
    if title in system[1::]:
        div2_2.pack(fill=BOTH,expand=True)
        if title == "left click":
            title_input_number.set("ใส่จำนวนคลิ๊ก : ")
            input_number.set("1")
        elif title == "move mouse":
            title_input_number.set("ใส่แกน (x,y) : ")
            input_number.set("100,100")
        elif title == "scroll mouse":
            title_input_number.set("ใส่ตัวเลขที่อยากให้เลื่อนเมาส์ (-)ลง,(+)ขั้น : ")
            input_number.set("50")
        elif title == "CTRL + key":
            title_input_number.set("CTRL + : ")
            input_number.set("c")
        elif title == "save url":
            title_input_number.set("ตั้งชื่อไฟล์ที่คุณต้องการบันทึก : ")
            input_number.set("ชื่อไฟล์")
        elif title == "ID":
            title_input_number.set("ใส่ตัวเลขตามบรรทัดใน NameUrlSave : ")
            input_number.set("0")
        else:
            form1.forget()
    else:
        div2_2.forget()
        if title == "time delay":
            input_number.set("1")
            title_input_number.set("ใส่วินาที ที่ต้องการดีเลย์เวลา : ")
        else:
            form1.forget()
        
# div1
label1 = ttk.Label(divPage1,text='เพิ่มคำสั่ง',font=font1)
label1.pack()
system = ["time delay","left click","right click","move mouse","scroll mouse","enter","CTRL + key","click20Item","check 0","save url","DownloadExcel","ID"]
for command in system:
    button = ttk.Button(divPage1, text=command, command=lambda cmd=command: setOption(cmd))
    button.pack()
divpositionmouse = Frame(divPage1).pack(pady=50)
ttk.Label(divpositionmouse,text="หาตำแหน่ง x,y เมาส์ ดีเลย์ 3 วิ",font=font2).pack()
ttk.Button(divpositionmouse,text="Position",command=positionMouse).pack()
ttk.Label(divpositionmouse,textvariable=ganXY,font=font2).pack()

# div2
ttk.Label(div2_1, text='รูปแบบคำสั่งที่เลือก', font=font1).pack()
ttk.Label(div2_1, textvariable=title_option, font=font2).pack()

form1 = Frame(div2_1)
ttk.Label(form1,textvariable=title_input_number,font=normal).pack(side=LEFT,expand=True)
ttk.Entry(form1,textvariable=input_number).pack(expand=True)

ttk.Label(div2_2,textvariable=testSys,font=font2).pack(side=TOP)
ttk.Button(div2_2,text="test",command=Runtest_system).pack(side=LEFT,fill=BOTH,expand=True,padx=20)
ttk.Button(div2_2,text="stop",command=stoptest_system).pack(side=LEFT,fill=BOTH,expand=True,padx=20)

addCommand =  ttk.Button(div2_3,text="add to commands",command=AddCommands_system)

##############  หน้า2 ################
def open_csv_file():
    table.delete(*table.get_children())
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    pathFile.set(file_path)
    if file_path:
        with open(file_path, newline='') as csv_file:
            csv_reader = csv.reader(csv_file)
            i = 0
            for row in csv_reader:
                if i!=0:
                    try:
                        if ':' in row[0]:
                            s1 = row[0].split(':')
                            if ',' in s1[1]:
                                numberState = s1[1].split(',')
                                table.insert('','end',values=(i,s1[0], f'x={numberState[0]},y={numberState[1]}'))
                            else:
                                table.insert('','end',values=(i,s1[0],s1[1]))
                    except:
                        pass
                i+=1

def deleteStep():
    try:
        id = int(idStep.get())
        df = pd.read_csv(pathFile.get())
        df = df.drop(id-1)
        df = df.reset_index(drop=True)
        df.to_csv(pathFile.get(), index=False)
        idStep.set("")
        setTreeCommand()
    except:
        messagebox.showerror("error","step ที่จะลบต้องเป็นตัวเลขเท่านั้นห้ามเป็นค่าว่างและต้องมี step ในตาราง")

def StartCommands():
    try:
            loopx = int(loop.get())
    except:
        loop.set(1)
        return
    for l in range(loopx):
        # print(f'L = {l+1}')
        with open(pathFile.get(),newline='') as commands:
            command = csv.reader(commands)
            i = 0
            for step in command:
                if i != 0:
                    # print(f'i = {i} , step = {step}s')
                    ReadStepRun.set(i)
                    if statusRunStartCommands:
                        return
                    text = []
                    try:
                        text = step[0].split(':')
                    except:
                        messagebox.showerror("error",f"คำสั่ง {step[0]} error ลบอกแล้วสร้างคำสั่งใหม่")
                        return
                    if text[0] == 'timesleep':
                        try:
                            time.sleep(int(text[1]))
                        except:
                            messagebox.showerror("error",f"คำสั่ง {text[0]} error {text[1]} ไม่อยู่ในรูปแบบตัวเลข")
                            return
                    elif text[0] == 'leftClick':
                        try:
                            for i1 in range(int(text[1])):
                                at.click()
                        except:
                            messagebox.showerror("error",f"คำสั่ง {text[0]} error {text[1]} ต้องเป็นตัวเลขระหว่าง 0-100")
                            return
                    elif text[0] == 'rightClick':
                        try:
                            for i1 in range(int(text[1])):
                                at.rightClick()
                        except:
                            messagebox.showerror("error",f"คำสั่ง {text[0]} error {text[1]} ต้องเป็นตัวเลขระหว่าง 0-100")
                            return
                    elif text[0] == 'moveTo':
                        try:
                            xy = text[1].split(',')
                            at.moveTo(int(xy[0]),int(xy[1]))
                        except:
                            messagebox.showerror("error",f"คำสั่ง {text[0]} error {text[1]} ไม่อยู่ในรูปแบบ x,y และต้องเป็นตัวเลขเท่านั้น")
                            return
                    elif text[0] == 'scroll':
                        try:
                            at.scroll(int(text[1]))
                        except:
                            messagebox.showerror("error",f"คำสั่ง {text[0]} error {text[1]} ไม่อยู่ในรูปแบบ x,y และต้องเป็นตัวเลขเท่านั้น")
                            return
                    elif text[0] == 'enter':
                        at.press('enter')
                    elif text[0] == 'CTRL':
                        try:
                            at.hotkey('ctrl',str(text[1]))
                        except:
                            messagebox.showerror("error",f"คำสั่ง {text[0]} error {text[1]} ไม่อยู่ในรูปแบบ x,y และต้องเป็นตัวเลขเท่านั้น")
                            return
                    elif text[0] == 'click20Item':
                        time.sleep(1)
                        web.select_item20()
                        time.sleep(1)
                    elif text[0] == 'check':
                        time.sleep(1)
                        web.checkZero()
                        time.sleep(1)
                    elif text[0] == 'save url':
                        time.sleep(1)
                        indexLoop = int(loop.get())
                        try:
                            web.saveUrl(name=ReadNameUrlSave()[l])
                        except:
                            web.saveUrl(name=(l+1))
                    elif text[0] == 'DownloadExcel':
                        time.sleep(1)
                        web.DownloadExcel()
                        time.sleep(1)
                    elif text[0] == 'ID':
                        try:
                            pyperclip.copy(ReadNameUrlSave()[l])
                        except:
                            pyperclip.copy(l)
                i+=1
            if i <=1 : messagebox.showerror("error","กรุณาเพิ่มคำสั่งในตาราง")

def RunStartCommands():
    global statusRunStartCommands
    statusRunStartCommands = False
    t = threading.Thread(target=StartCommands)
    t.start()

def StopStartCommands():
    global statusRunStartCommands
    statusRunStartCommands = True
    messagebox.showinfo("หยุดทำงาน step",f"หยุดทำงาน step")

loop.set(1)
ttk.Button(t2, text="select Log", command=setTreeCommand).place(x=10,y=10)
ttk.Button(t2, text="open fils.csv", command=open_csv_file).place(x=10,y=50)
ttk.Label(t2,text="จำนวนรอบ :",font=normal).place(x=10,y=85)
ttk.Entry(t2,textvariable=loop,width=5).place(x=100,y=90)
ttk.Button(t2, text="start", command=RunStartCommands).place(x=10,y=130)
ttk.Button(t2, text="stop", command=StopStartCommands).place(x=10,y=160)

ttk.Label(t2,text="ลบ step",font=font2).place(x=10,y=200)
ttk.Label(t2,text="input step :",font=normal).place(x=10,y=240)
ttk.Entry(t2,textvariable=idStep,width=5).place(x=85,y=250)
ttk.Button(t2, text="delete", command=deleteStep).place(x=10,y=280)

ReadStepRun.set(0)
ttk.Label(t2,text="ทำงาน step :",font=normal).place(x=10,y=350)
ttk.Label(t2,textvariable=ReadStepRun,font=normal).place(x=100,y=350)

header = ['step','System','about']
hdsize = [50,200,200]
table = ttk.Treeview(t2,columns=header,show='headings')
table.place(x=150,y=10,height=430)

# header
for h,s in zip(header,hdsize):
    table.heading(h,text=h)
    table.column(h,width=s)
setTreeCommand()

##############  หน้า3 ################
def setElementPageXPathAll(path):
    listpath = path.split('/')
    index = 0
    for i in listpath:
        try:
            if int(i[-2]):
                title = i[0:-3]+f':nth-of-type({i[-2]})'
                listpath[index] = title
        except:
            pass
        index+=1
    resetPath = ''
    for i in listpath:
        if len(i)>0:
            resetPath += i+'>'
    return resetPath[0:-1]

urllink = StringVar()
extension_path1 = './LogPrograme/3.0.8_0.crx'
extension_path2 = './LogPrograme/13.0.10_0.crx'

def setElementPageXPathAll(path):
    listpath = path.split('/')
    index = 0
    for i in listpath:
        try:
            if int(i[-2]):
                title = i[0:-3]+f':nth-of-type({i[-2]})'
                listpath[index] = title
        except:
            pass
        index+=1
    resetPath = ''
    for i in listpath:
        if len(i)>0:
            resetPath += i+'>'
    return resetPath[0:-1]

class webdriverGet():
    def __init__(self):
        try:
            self.web_options = webdriver.ChromeOptions()
            self.web_options.add_extension(extension_path1)
            self.web_options.add_extension(extension_path2)
            self.web_options.add_experimental_option("prefs", {
                "download.prompt_for_download": True,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True
            })
        except:
            return {"message":"error set Webdriver","status":404}
        
    def open(self):
        try:
            self.driver = webdriver.Chrome(options=self.web_options)
        except:
            return {"message":"error open webdriver","status":404}
            
    def ToLink(self,url):
        try:
            self.driver.switch_to.window(self.driver.window_handles[0])
            self.driver.get(url=url)
        except:
            return {"message":"error get Url","status":404}
        
    def close(self):
        try:
            self.driver.quit()
        except:
            return {"message":"error close webdriver","status":404}

    def checkZero(self):
        try:
            # select 0 
            self.driver.switch_to.window(self.driver.window_handles[-1])
            self.data = self.driver.page_source
            self.soup = bs4.BeautifulSoup(self.data)
            for i in range(20,0,-1):
                time.sleep(2)
                self.get_number = self.soup.select_one(f'body > div > div > div > div.ap-page__bd > div > div.ap-compare-products__bd > div > table > tbody:nth-child(3) > tr:nth-child(3) > td:nth-child({i+1}) > div')
                try:
                    if int(self.get_number.text) == 0:
                        self.driver.find_element(By.XPATH,f"/html/body/div/div/div/div[2]/div/div[2]/div/table/thead/tr/th[{i+1}]/div[3]").click()
                except:
                    return {"message":"ไม่พบสินค้า...","status":200}
            return {"message":"เช็ค 0 สินค้าสำเร็จ","status":200}
        except:
            return {"message":"ไม่พบหน้าที่ต้องการให้เช็ค 0","status":404}
        
    def select_item20(self):
        try:
            # set web index 0
            self.driver.switch_to.window(self.driver.window_handles[0])
            # click 20 item
            for index in range(20):
                time.sleep(1)
                self.driver.find_element(By.XPATH,f'//*[@id="ap-sbi-alibabaCN-result"]/div/div[2]/div/div[1]/div[2]/div/div/div/div/div[{index+1}]/div/div[10]').click()
                self.element = self.driver.find_element(By.XPATH, f'//*[@id="ap-sbi-alibabaCN-result"]/div/div[2]/div/div[1]/div[2]/div/div/div/div/div[{index+2}]')
                self.driver.execute_script("arguments[0].scrollIntoView();", self.element)
            time.sleep(1)
        except:
            return {"message":"error select item","status":404}

    def DownloadExcel(self):
        try:
            # download excel
            self.driver.find_element(By.XPATH,'/html/body/div/div/div/div[2]/div/div[1]/div/div').click()
        except:
            return {"message":"error Download file","status":404}
    
    def saveUrl(self,name):
        try:
            self.driver.switch_to.window(self.driver.window_handles[-1])
            open(f'{name}.txt','w',encoding='utf-8').write(self.driver.current_url)
        except:pass

web = webdriverGet()

def startWebsite():
    web.open()

def getUrl():
    try:
        web.ToLink(url=str(urllink.get()))
    except:pass

def click20Item():
    try:
        web.select_item20()
    except:pass

def check0():
    try:
        ss = web.checkZero()
        if(ss['status'] == 200):
            messagebox.showinfo("success",ss["message"])
        else:
            messagebox.showinfo("error","กรุณา start webdriver ของโปรแกม")
    except:
        messagebox.showinfo("error","กรุณา start webdriver ของโปรแกม")

def LoadExcelFile():
    try:
        web.DownloadExcel()
    except:pass

ttk.Button(t3,text="start webdriver",command=startWebsite).pack(pady=20)
ttk.Label(t3,text="URL:",font=font2).place(x=10,y=50)
ttk.Entry(t3,textvariable=urllink).place(x=65,y=63,width=500)
ttk.Button(t3,text="go to URL",command=getUrl).place(x=570,y=62)
window.mainloop()