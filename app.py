from tkinter import*
from tkinter import ttk
from tkinter import messagebox,filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys 
from PIL import ImageGrab
from datetime import datetime
import os,csv,pandas as pd, numpy as np, pyautogui as at
import requests,json,time,bs4,hashlib,threading

extension_path1 = './LogPrograme/3.0.8_0.crx'
extension_path2 = './LogPrograme/13.0.10_0.crx'
web_options = webdriver.ChromeOptions()
web_options.add_extension(extension_path1)
web_options.add_extension(extension_path2)



programname = "Bot 1688"
window = Tk()
window.title(programname)
window.config(bg="#ffffff") 
window.geometry("500x800-0+0")
tabs = ttk.Notebook(window)
tab1 = Frame(tabs)
tab2 = Frame(tabs)
tabs.pack(fill=BOTH,expand=1)
tabs.add(tab1,text="home")
tabs.add(tab2,text="settings")

normal = ('Angsana New',15)
font1 = ('Angsana New',25)
font2 = ('Angsana New',20)

# //////////////////////////////////////////// system programe setting ///////////////////////////////////////////////////////////////////////
def setLog1688(newSetting):
    datas = open("./dataProgram/log1688.txt",'w',encoding="utf8")
    datas.write(newSetting)
    datas.close()
def getSettingsProgram():
    datas = open("./dataProgram/setting.txt",'r',encoding="utf8")
    result = {}
    for data in datas.readlines():
        splitdatas = data.split(',,,')
        for i in splitdatas:
            j = i.split('=')
            if(len(j)==2):
                if(j[0] == "positionDowload" or j[0] == "position_img"):
                    g = j[1].split(',')
                    result[j[0]] = (int(g[0]),int(g[1]))
                else:
                    result[j[0]] = j[1]
    return result
def setSettingsProgram(newSetting):
    data = ""
    for k in newSetting:
        if(k == "positionDowload" or k == "position_img"):
            replaceValueTuple = str(newSetting[k]).replace('(','').replace(')','').replace(' ','')
            data += f"{k}={replaceValueTuple},,,"
        else:
            data += f"{k}={newSetting[k]},,,"
    datas = open("./dataProgram/setting.txt",'w',encoding="utf-8")
    datas.write(data)
    datas.close()
def selectPathDownload():
    folder_selected = filedialog.askdirectory()
    if(len(folder_selected)!=0):
        npath = getSettingsProgram()
        npath["pathdownload"] = folder_selected
        titlePath.set(folder_selected)
        setSettingsProgram(npath)
def selectPathAPI():
    NewpathAPI = titleAPI.get( )
    npath = getSettingsProgram()
    npath["urlAPI"] = NewpathAPI
    setSettingsProgram(npath)
def selectDownCout():
    NewpathDownCount = titleDownCout.get( )
    npath = getSettingsProgram()
    npath["downcount"] = NewpathDownCount
    setSettingsProgram(npath)
def selecttimeSleep20items():
    Newseting = titletimeSleep20items.get()
    npath = getSettingsProgram()
    npath["timesleep20item"] = Newseting
    setSettingsProgram(npath)
def selectPositionDowload():
    time.sleep(2)
    Newposition = at.position()
    titlePositionDowload.set(f"({Newposition.x},{Newposition.y})")
    npath = getSettingsProgram()
    npath["positionDowload"] = f"{Newposition.x},{Newposition.y}"
    setSettingsProgram(npath)
# //////////////////////////////////////////// system webdriver ///////////////////////////////////////////////////////////////////////////
groups = {
    'keyvalue':{
        "อุปกรณ์เสริม-อิเล็กทรอนิกส์":{"thai":"อุปกรณ์เสริม-อิเล็กทรอนิกส์","eng":"group1"},
        "อุปกรณ์-อิเล็กทรอนิกส์":{"thai":"อุปกรณ์-อิเล็กทรอนิกส์","eng":"group2"},
        "ทีวีและเครื่องใช้ในบ้าน":{"thai":"ทีวีและเครื่องใช้ในบ้าน","eng":"group3"},
        "สุขภาพและความงาม":{"thai":"สุขภาพและความงาม","eng":"group4"},
        "ทารกและของเล่น":{"thai":"ทารกและของเล่น","eng":"group5"},
        "ของชำและสัตว์เลี้ยง":{"thai":"ของชำและสัตว์เลี้ยง","eng":"group6"},
        "บ้านและไลฟ์สไตล์":{"thai":"บ้านและไลฟ์สไตล์","eng":"group7"},
        "แฟชั่นและเครื่องประดับผู้หญิง":{"thai":"แฟชั่นและเครื่องประดับผู้หญิง","eng":"group8"},
        "แฟชั่นและเครื่องประดับผู้ชาย":{"thai":"แฟชั่นและเครื่องประดับผู้ชาย","eng":"group9"},
        "กีฬาและการเดินทาง":{"thai":"กีฬาและการเดินทาง","eng":"group10"},
        "ยานยนต์และรถจักรยานยนต์":{"thai":"ยานยนต์และรถจักรยานยนต์","eng":"group11"},
        "เครื่อเขียนหนังสือ":{"thai":"เครื่อเขียนหนังสือ","eng":"group12"},
        "เครื่องประดับ":{"thai":"เครื่องประดับ","eng":"group13"},
        "ตั๋วและบัตรกำนัน":{"thai":"ตั๋วและบัตรกำนัน","eng":"group14"},
    },
    'array':["อุปกรณ์เสริม-อิเล็กทรอนิกส์","อุปกรณ์-อิเล็กทรอนิกส์","ทีวีและเครื่องใช้ในบ้าน","สุขภาพและความงาม","ทารกและของเล่น","ของชำและสัตว์เลี้ยง","บ้านและไลฟ์สไตล์",
             "แฟชั่นและเครื่องประดับผู้หญิง","แฟชั่นและเครื่องประดับผู้ชาย","กีฬาและการเดินทาง","ยานยนต์และรถจักรยานยนต์","เครื่อเขียนหนังสือ","เครื่องประดับ","ตั๋วและบัตรกำนัน"]
}

CNY_TO_THB = 5.30
url = "https://api.exchangerate-api.com/v4/latest/CNY"
response = requests.get(url)
if response.status_code == 200:
    data = response.json()
    CNY_TO_THB = data["rates"]["THB"]

def sha256(data):
    return hashlib.sha256(data.encode()).hexdigest()
def readCsvAll(folder_path='./'):
    csv_files = []
    try:
        for filename in os.listdir(folder_path):
            if filename.endswith('.csv'):
                csv_files.append(filename)
    except:pass
    return csv_files
def readFileCsv(name):
    dataCsv = []
    # try:
    with open(name, mode='r', newline='', encoding='utf-8') as csv_file:
        csv_reader = csv.reader(csv_file)
        for row in csv_reader:
            dataCsv.append(row)
    # except:pass
    return dataCsv
def postDataAPI1688(data,group,idproduct):
    uri = f"{getSettingsProgram()["urlAPI"]}postdataby1688?key=5a3dec84301206e275f7ca7fa119796c8a5be05d100a2d23ba3a4f189876d03a"
    res = requests.post(uri,data={
        "item20":data,
        "group":group,
        "idproduct":idproduct
    })
    return res
def formatHeader(i):
    key = "";
    if("ชื่อร้าน" in i):
        key = "namestore"
    elif("ชื่อผลิตภัณฑ์" in i):
        key = "nameproduct"
    elif("ราคา" in i):
        key = "priceCNY"
    elif("ลิงค์ผลิตภัณฑ์" in i):
        key = "linkproduct"
    elif("ลิงค์รูปภาพ" in i):
        key = "linkimage"
    elif("จำนวนคำสั่งชำระเงินใน 30 วันที่ผ่านมา" in i):
        key = "countpayment30daylater"
    elif("อัตราการฟื้นตัว 48H ในช่วง 30 วันที่ผ่านมา" in i):
        key = "lateheal30daylater"
    elif("อัตราประสิทธิภาพ 48H ในช่วง 30 วันที่ผ่านมา" in i):
        key = "efficient30daylater"
    elif("อัตราการตอบกลับ 3 นาทีในช่วง 30 วันที่ผ่านมา" in i):
        key = "reply3minute30daylater"
    elif("อัตราการคืนเงินที่มีคุณภาพใน 30 วันที่ผ่านมา" in i):
        key = "refund30daylater"
    elif("อัตราการโต้แย้งใน 30 วันที่ผ่านมา" in i):
        key = "argue30daylater"
    elif("การรับรองความแข็งแกร่ง 1688" in i):
        key = "late1688"
    elif("บริการทั่วไป" in i):
        key = "aboutsevice"
    elif("อัตราผลตอบแทน" in i):
        key = "reward"
    elif("แลกเปลี่ยนประสบการณ์" in i):
        key = "experience"
    elif("ประสบการณ์ที่มีคุณภาพ" in i):
        key = "quality"
    elif("ความตรงต่อเวลาของโลจิสติกส์" in i):
        key = "deliveredtime"
    elif("การระงับข้อพิพาท" in i):
        key = "A1"
    elif("ที่ปรึกษาการจัดซื้อจัดจ้าง" in i):
        key = "A2"
    elif("ลักษณะการลงทะเบียน" in i):
        key = "A3"
    elif("ช่วงเข้า" in i):
        key = "AM"
    elif("รูปแบบธุรกิจ" in i):
        key = "typebusiness"
    elif("ที่ตั้ง" in i):
        key = "addresshere"
    elif("ที่อยู่" in i):
        key = "addresslive"
    elif("ข้อมูลติดต่อ" in i):
        key = "teldata"
    elif("ระยะขอบ" in i):
        key = "margin"
    elif("ที่จัดตั้งขึ้น" in i):
        key = "esabished"
    elif("พื้นที่องค์กร" in i):
        key = "areaorganization"
    elif("จำนวนพนักงาน" in i):
        key = "countemployee"
    else:
        key = i
    return key
def saveScreenShot(pathFile):
    im = ImageGrab.grab()
    im.save(pathFile)
def sendMessageAndImageLine(title,pathFile):
    now = datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H-%M-%S")
    LINE_ACCESS_TOKEN="kQlzp9toM8Js213IiJCys6qV4cwOuXcXFmp6UcwzgmK"
    url = "https://notify-api.line.me/api/notify"
    file = {'imageFile':open(pathFile,'rb')}
    data = ({
            'message':f'หยุดทำงาน\n{title}{date_str}{time_str}'
        })
    LINE_HEADERS = {"Authorization":"Bearer "+LINE_ACCESS_TOKEN}
    session = requests.Session()
    r=session.post(url, headers=LINE_HEADERS, files=file, data=data)
def sendMessageLine(title):
    now = datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H-%M-%S")
    LINE_ACCESS_TOKEN="kQlzp9toM8Js213IiJCys6qV4cwOuXcXFmp6UcwzgmK"
    url = "https://notify-api.line.me/api/notify"
    data = ({
        'message':f'สำเร็จ\n{title} ของ {date_str}{time_str}'
    })
    LINE_HEADERS = {"Authorization":"Bearer "+LINE_ACCESS_TOKEN}
    session = requests.Session()
    r=session.post(url, headers=LINE_HEADERS, data=data)
def WarningError(title):
    now = datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H-%M-%S")
    pathFile = './imageWarning/{}{}{}.png'.format(title,date_str,time_str)
    saveScreenShot(pathFile)
    sendMessageAndImageLine(title,pathFile)
# ///////////////////////////////////////////////
download_path1 = getSettingsProgram()["pathdownload"]
download_path = download_path1.replace("/","\\")
def systemSendAPI(dataInLog):
    filesCSV = readCsvAll(download_path1)
    try:
        df = readFileCsv(download_path1+'/'+filesCSV[-1])
        header = df[0]
        sumproduct = []
        sumid = {}
        for datacsv in df[1::]:
            obj = {}
            for row in range(len(header)):
                obj[formatHeader(header[row])] =  datacsv[row]
            obj["priceCNY"] = float(obj["priceCNY"].split('¥')[1])
            obj["priceTHB"] = float(CNY_TO_THB*float(obj["priceCNY"]))
            obj["id"] = sha256(obj['linkproduct'])
            sumid[sha256(obj['linkproduct'])] = obj
            sumproduct.append(obj)
        postDataAPI1688(json.dumps(sumid),dataInLog["groups"],dataInLog["id"])
        return 200
    except:
        return 404
def fetchAPI(selectGroup,page):
    key = "382557d3731870efb6e863c174ccb351f191987e79a9c5805e16698e055b4364";
    uri = f"{getSettingsProgram()["urlAPI"]}getdatabybot1688?key={key}&&group={selectGroup}&&page={page}"
    try:
        return requests.get(uri);
    except:
        return 404
titleSelectGroup = StringVar()
titleSelectGroup.set(getSettingsProgram()["selectGroup"])
# 
titleSelectPage = StringVar()
titleSelectPage.set(getSettingsProgram()["selectPage"])
totalPages = StringVar()
totalPages.set(getSettingsProgram()["totalPages"])
# 
titleSelectNumber = StringVar()
titleSelectNumber.set(getSettingsProgram()["SelectNumber"])
totalNumber = StringVar()
totalNumber.set(getSettingsProgram()["totalNumber"])
# 
titleShowStatus = StringVar()
titleShowStatus.set("หยุดทำงาน")

web_options.add_experimental_option("prefs", {
    "download.default_directory": download_path,
})
driver = webdriver.Chrome(options=web_options)
driver.get('https://www.aliprice.com/Member/login.html?ext_id=10100&platform=1688&version=3.0.8&browser=chrome&channel=chrome&mv=3&redirect=https%3A%2F%2Fwww.1688.com')
driver.switch_to.window(driver.window_handles[-1])
driver.execute_script("window.close();")
time.sleep(1)
driver.switch_to.window(driver.window_handles[-1])
driver.execute_script("window.close();")
driver.switch_to.window(driver.window_handles[0])
# Login 1688
# email : tmhumr1frz2vjv7nf1i62lc@sharklasers.com
# password : passapol47
urlserver = 'http://localhost:3030/'
def setSelectPage():
    newString = titleSelectPage.get()
    npath = getSettingsProgram()
    npath["selectPage"] = newString
    setSettingsProgram(npath)
def setSelectNumber():
    newString = titleSelectNumber.get()
    npath = getSettingsProgram()
    npath["SelectNumber"] = newString
    setSettingsProgram(npath)
def setSelectTotalPage():
    newString = totalPages.get()
    npath = getSettingsProgram()
    npath["totalPages"] = newString
    setSettingsProgram(npath)
def setSelectTotalNumber():
    newString = totalNumber.get()
    npath = getSettingsProgram()
    npath["totalNumber"] = newString
    setSettingsProgram(npath)
def setTitleSelectGroup(event):
    titleSelectGroup.set(event.widget.get())
    npath = getSettingsProgram()
    npath["selectGroup"] = event.widget.get()
    setSettingsProgram(npath)

def startBot():
    print(titleShowStatus.get())
    setSelectPage()
    setSelectNumber()
    selectGroup = getSettingsProgram()["selectGroup"]
    selectPage = int(getSettingsProgram()["selectPage"])
    selctNumber = int(getSettingsProgram()["SelectNumber"])-1
    if(selctNumber<=0):
        selctNumber=1
    res = fetchAPI(selectGroup,selectPage)
    if(res==404):
        WarningError(f"ดึงข้อมูล {selectGroup} จากเชิฟเวอร์APIไม่ได้")
        titleShowStatus.set("หยุดทำงาน")
        return
    else:
        res = json.loads(fetchAPI(selectGroup,selectPage).text)
        while(True):
            if(titleShowStatus.get()=="หยุดทำงาน"):
                return
            # docs totalPages hasNextPage page
            dataAPI = res["docs"]
            totalPages.set(res["totalPages"])
            totalNumber.set(len(dataAPI))
            setSelectTotalPage()
            setSelectTotalNumber()
            # print(len(res["docs"]))
            for count in range(selctNumber,len(dataAPI)):
                if(titleShowStatus.get()=="หยุดทำงาน"):
                    return
    # ////////////////////////////////// set product//////////////////////////////////
                dataInLog = {
                    "groups":dataAPI[count]['group'],
                    "page":res["page"],
                    "number":count+1,
                    "id":dataAPI[count]['idproduct'],
                    "images":dataAPI[count]['image_product_1'],
                }
                # print(len(res["docs"]))
                # print(dataInLog)
    # ////////////////////////////////// get image //////////////////////////////////
                if(titleShowStatus.get()=="หยุดทำงาน"):
                    return
                try:
                    driver.get(f'{urlserver}?url={dataInLog["images"]}')
                except:
                    WarningError("เซิฟเวอร์แสดงรูปของบอท 1688 มีปัญหา")
                    titleShowStatus.set("หยุดทำงาน")
                    return
                time.sleep(int(getSettingsProgram()["timesleepLoadimage"]))
    # ////////////////////////////////// click sheach //////////////////////////////////
                if(titleShowStatus.get()=="หยุดทำงาน"):
                    return
                at.click(getSettingsProgram()["position_img"])
                at.rightClick()
                # time.sleep(0.5) 
                for i in range(int(getSettingsProgram()["downcount"])):
                    at.press('down')
                at.press('enter')
                at.press('enter')
                time.sleep(int(getSettingsProgram()["timesleepLoadSheach"]))
    # ////////////////////////////////// //////////////////////////////////////////////
                try:
                    for index in range(20):
                        if(titleShowStatus.get()=="หยุดทำงาน"):
                            return
                        else:
                            try:
                                driver.find_element(By.XPATH,f'//*[@id="ap-sbi-alibabaCN-result"]/div/div[2]/div/div[1]/div[2]/div/div/div/div/div[{index+1}]/div/div[10]').click()
                            except:
                                driver.find_element(By.XPATH,f'//*[@id="ap-sbi-alibabaCN-result"]/div/div[2]/div/div[1]/div[2]/div/div/div/div/div[{index+1}]/div/div[9]').click()
                            element = driver.find_element(By.XPATH, f'//*[@id="ap-sbi-alibabaCN-result"]/div/div[2]/div/div[1]/div[2]/div/div/div/div/div[{index+2}]')
                            driver.execute_script("arguments[0].scrollIntoView();", element)
                    driver.find_element(By.XPATH,f'/html/body/div[2]/div/div[1]/div[2]/div/div[2]').click() 
                except:
                    WarningError("เว็บไซด์มีปัญหา ติ๊กสินค้าในแว่นตาไม่ได้")
                    titleShowStatus.set("หยุดทำงาน")
                    return
                if(titleShowStatus.get()=="หยุดทำงาน"):
                    return
                time.sleep(int(getSettingsProgram()["timesleep20item"]));
    # ////////////////////////////////// click download CSV //////////////////////////////////
                if(titleShowStatus.get()=="หยุดทำงาน"):
                    return
                # at.position()
                at.click(getSettingsProgram()["positionDowload"])
                time.sleep(int(getSettingsProgram()["timesleepReadCSV"]));
    # ////////////////////////////////// system send API //////////////////////////////////
                if(titleShowStatus.get()=="หยุดทำงาน"):
                    return
                statusSendAPI = systemSendAPI(dataInLog)
                if(statusSendAPI == 404):
                    WarningError("หาไฟล์ไฟล์สินค้าไม่เจอ อาจจะเกิดจากโหลดไม่สำเร็จหรือปุ่มกดโหลด excel ไม่ถูกที่")
                    titleShowStatus.set("หยุดทำงาน")
                    return
    # ////////////////////////////////// save log And send line //////////////////////////////////
                setLog1688(str(dataInLog))
                sendMessageLine(f"{dataInLog["groups"]} Page: {dataInLog["page"]}/{res["totalPages"]} ชิ้นที่ {dataInLog["number"]}/{len(dataAPI)}")
    # ////////////////////////////////// close reset bowser//////////////////////////////////
                driver.switch_to.window(driver.window_handles[-1])
                driver.execute_script("window.close();")
                driver.switch_to.window(driver.window_handles[0])
    # ////////////////////////////////// //////////////////////////////////
            if(titleShowStatus.get()=="หยุดทำงาน"):
                return
            if(selctNumber < len(dataAPI)):
                titleSelectNumber.set(int(titleSelectNumber.get())+1)
            if(res["hasNextPage"]):
                selectPage += 1
                titleSelectPage.set(int(titleSelectPage.get())+1)
                titleSelectNumber.set(1)
                res = json.loads(fetchAPI(selectGroup,selectPage).text)
            else:
                break
def start():
    titleShowStatus.set("กำลังทำงาน")
    t = threading.Thread(target=startBot)
    t.start()

def stop():
    titleShowStatus.set("หยุดทำงาน")
# //////////////////////////////////////////// tab1 ///////////////////////////////////////////////////////////////////////
dropdown = ttk.Combobox(tab1, textvariable=titleSelectGroup,values=groups["array"]);
dropdown.place(x=10,y=10);
dropdown.bind("<<ComboboxSelected>>", setTitleSelectGroup)

LableSelectPage = Label(tab1,text="Page :")
LableSelectPage.place(x=10,y=40)
inputSelectPage = Entry(tab1, textvariable=titleSelectPage)
inputSelectPage.place(x=100,y=40);
LableTotalPage = Label(tab1,textvariable=totalPages)
LableTotalPage.place(x=300,y=40)

LableSelectPage = Label(tab1,text="Number :")
LableSelectPage.place(x=10,y=80)
inputSelectPage = Entry(tab1, textvariable=titleSelectNumber)
inputSelectPage.place(x=100,y=80);
LableTotalNumber = Label(tab1,textvariable=totalNumber)
LableTotalNumber.place(x=300,y=80)


LableShowStatus = Label(tab1,textvariable=titleShowStatus)
LableShowStatus.place(x=10,y=180)
btnselectPath = Button(tab1,text="start",command=start)
btnselectPath.place(x=10,y=200)
btnselectPath = Button(tab1,text="stop",command=stop)
btnselectPath.place(x=60,y=200)
# //////////////////////////////////////////// tab2 ///////////////////////////////////////////////////////////////////////
titlePath = StringVar()
titlePath.set(getSettingsProgram()["pathdownload"])
# 
titleAPI = StringVar()
titleAPI.set(getSettingsProgram()["urlAPI"])
# 
titleDownCout = StringVar()
titleDownCout.set(getSettingsProgram()["downcount"])
# 
titletimeSleep20items = StringVar()
titletimeSleep20items.set(getSettingsProgram()["timesleep20item"])
# 
titlePositionDowload = StringVar()
titlePositionDowload.set(str(getSettingsProgram()["positionDowload"]))
# ////////////////////////////////////////////
btnselectPath = Button(tab2,text="Path",command=selectPathDownload)
btnselectPath.place(x=10,y=20)
LabletitlePath = Label(tab2,textvariable=titlePath)
LabletitlePath.place(x=60,y=20)
# ////////////////////////////////////////////
btnUriAPI = Button(tab2,text="API",command=selectPathAPI)
btnUriAPI.place(x=10,y=60)
inputUriAPI = Entry(tab2,textvariable=titleAPI)
inputUriAPI.place(x=60,y=60)
# ////////////////////////////////////////////
btnDowncount = Button(tab2,text="DownCout",command=selectDownCout)
btnDowncount.place(x=10,y=100)
inputDowncount = Entry(tab2,textvariable=titleDownCout)
inputDowncount.place(x=110,y=105)
# ////////////////////////////////////////////
btntimeSleep20items = Button(tab2,text="time20Item",command=selecttimeSleep20items)
btntimeSleep20items.place(x=10,y=140)
inputtimeSleep20items = Entry(tab2,textvariable=titletimeSleep20items)
inputtimeSleep20items.place(x=110,y=145)
# ////////////////////////////////////////////
btnPositionDowload = Button(tab2,text="clickDownload",command=selectPositionDowload)
btnPositionDowload.place(x=10,y=180)
LabletitlePositionDowload = Label(tab2,textvariable=titlePositionDowload)
LabletitlePositionDowload.place(x=130,y=185)


window.mainloop()
