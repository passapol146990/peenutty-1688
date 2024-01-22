import os,csv,pandas as pd,numpy as np

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
    try:
        with open(name, mode='r', newline='', encoding='utf-8') as csv_file:
            csv_reader = csv.reader(csv_file)
            for row in csv_reader:
                dataCsv.append(row)
    except:pass
    return dataCsv
#////////////////////////////////////////////////////////////////
try:
    df = pd.read_excel('./dataHead.xlsx')
    header = df.columns
    DataFileExcel = [] 
    for idx in range(len(df.values)):
        values = df.values[idx]
        DataFileExcel.append(values)
#////////////////////////////////////////////////////////////////
    DataResult = []
    for dataRead in DataFileExcel:
        dataStart = dataRead 
        for r in range(3):
            dataStart = np.append(dataStart,'')
        id = dataStart[0]
        s = readFileCsv(f'{id}.csv')
        if s:
            for idx in range(23):
                if idx == 0:
                    dataStart = np.append(dataStart,s[1::][idx])
                    DataResult.append(dataStart)
                else:
                    list1 = ['' for i in range(12)]
                    if(idx < len(s[1::])):
                        list1 = np.append(list1, s[1::][idx])
                    DataResult.append(list1)
#////////////////////////////////////////////////////////////////
    dataTest = pd.DataFrame(DataResult)
    dataTest.to_excel('./ResultData.xlsx')
except:
    open('./dataHead.xlsx', 'w',encoding='utf8')