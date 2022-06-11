import json
import requests
import htmls
from bs4 import BeautifulSoup
from lxml import etree
import time
import pandas
import datetime
import re

doneList = set()

columnName = ['department', 
              'tenderNum', 
              'tenderName', 
              'tenderMethod', 
              'tenderType', 
              'annouceDate', 
              'deadline', 
              'money', 
              'trendHref']
chineseColumnName = {
    'department' : '機關名稱', 
  'tenderNum' : '檔案案號', 
  'tenderName' : '標案名稱', 
  'tenderMethod' : '招標方式', 
  'tenderType' : '採購性質', 
  'annouceDate' : '公告日期', 
  'deadline' : '截止投標', 
  'money' : '預算金額', 
  'trendHref' : '連結'
}
inv_chineseColumnName = {v: k for k, v in chineseColumnName.items()}
outputColumn = dict()
for name in columnName:
    outputColumn[chineseColumnName[name]] = []

def createPostData(orgId, departmentName, StartTime, EndTime):
    postForm = {
        'pageSize':'',
        'firstSearch':'true',
        'searchType':'basic',
        'level_1':'on',
        'orgName':departmentName,
        'orgId':orgId,
        'tenderName':'',
        'tenderId':'',
        'tenderType':'TENDER_DECLARATION',
        'tenderWay':'TENDER_WAY_ALL_DECLARATION',
        'dateType':'isDate',
        'tenderStartDate':StartTime,
        'tenderEndDate':EndTime,
        'radProctrgCate':''
    }
    return postForm

def getColumnInf(column, var):
    return {
        'department' : column[1].text.replace('\t','').replace('\n','').replace('\r','').replace(' ',''),
        'tenderNum' : column[2].text.split("\r\n")[1].replace('\t',''),
        'tenderName' : re.findall('\("([^"]+)"\)', column[2].find("a").find("u").find("span").string)[0],
        'tenderMethod' : column[4].text.replace('\t','').replace('\n','').replace('\r','').replace(' ',''),
        'tenderType' : column[5].text.replace('\t','').replace('\n','').replace('\r','').replace(' ',''),
        'annouceDate' : column[6].text.replace('\t','').replace('\n','').replace('\r','').replace(' ',''),
        'deadline' : column[7].text.replace('\t','').replace('\n','').replace('\r','').replace(' ',''),
        'money' : column[8].text.replace('\t','').replace('\n','').replace('\r','').replace(' ',''),
        'trendHref' : "https://web.pcc.gov.tw{i}\n".format(i=column[2].find("a").get("href")),
    }.get(var,'error')

def getException():
    exceptionList = dict()
    try:
        with open('ExceptionList.json', encoding="utf-8") as f:
            data = json.load(f)
        exceptionList['department'] = data['機關名稱']
        exceptionList['tenderMethod'] = data['招標方式']
        exceptionList['tenderType'] = data['採購性質']
    except:
        print('[ERROR] Get Exception List False')
        exceptionList['department'] = []
        exceptionList['tenderMethod'] = []
        exceptionList['tenderType'] = []
    return exceptionList

basePath = "https://web.pcc.gov.tw"
buyPostPath = "https://web.pcc.gov.tw/prkms/tender/common/basic/readTenderBasic?"

if __name__ == '__main__':
    
    dayRange = input("How long do you want to colloect?(day) ")
    
    exceptionList = getException()
    
    currentTime = datetime.datetime.now()
    weekAgoTime = currentTime - datetime.timedelta(days = int(dayRange))
    StartTime = "{Y}/{M}/{D}".format(Y=weekAgoTime.year, M=weekAgoTime.month, D=weekAgoTime.day)
    StartTime = "{Y}/{MD}".format(Y=weekAgoTime.year, MD=weekAgoTime.strftime("%m/%d"))
    EndTime = "{Y}/{MD}".format(Y=currentTime.year, MD=currentTime.strftime("%m/%d"))
    
    print("Start Time: "+StartTime)
    print("End Time: "+EndTime)
    
    root = requests.get("https://web.pcc.gov.tw/prkms/tender/common/orgName/indexTenderOrgName")
    rootContent = root.content.decode()
    rootHtmltml = etree.HTML(rootContent)
    departments = rootHtmltml.xpath("//p/a[@style='color: blue;']")
    print(len(departments))
    for department in departments:
        print(department.attrib['title'])
        if department.attrib['title'] == '各級學校':
            print('Stop at 各級學校')
            break
        
        targetPath = basePath + department.attrib['href']
        targetRoot = requests.get(targetPath)
        targetSoup = BeautifulSoup(targetRoot.text, "html.parser")
        buyList = targetSoup.find("table", align='center', width='100%', border='0').findChildren("tr")
        for index in range(2, len(buyList)-1):
            itemColumn = buyList[index].findChildren('td')
            if len(itemColumn)<2:
                pass
            else:
                orgId = itemColumn[0].string
                departmentName = itemColumn[1].string
                postData = createPostData(orgId, departmentName, StartTime, EndTime)
                buy = requests.post(buyPostPath, data=postData)
                time.sleep(1) # 

                buySoup = BeautifulSoup(buy.text, "html.parser")
                items = buySoup.find("table", attrs={'class':'tb_01'}).find("tbody").findChildren("tr")
                column = None
                if orgId == None:
                    print('\t'+departmentName+' '+str(orgId)+' => None')
                elif len(items)<2:
                    print('\t'+departmentName+' '+str(orgId)+' => None')
                elif orgId in doneList:
                    print('\t'+departmentName+' '+str(orgId)+' => Done Before')
                else:
                    print('\t'+departmentName+' '+str(orgId)+' => '+str(len(items)))
                    
                    for item in items:
                        
                        isPass = False
                        infs = []
                        
                        # get trend inf
                        column = item.find_all("td")
                        try:
                            for name in columnName:
                                if not isPass:
                                    var = getColumnInf(column, name)
                                    infs.append(var)
                                    if name in exceptionList.keys():
                                        if var in exceptionList[name]:
                                            isPass = True
                        except:
                            print("Some Error!\n")
                        
                        if not isPass:
                            for i, name in enumerate(columnName):
                                outputColumn[chineseColumnName[name]].append(infs[i])
                    doneList.add(orgId)

    df = pandas.DataFrame(outputColumn)
    df.to_excel("國家機關採購案.xlsx",index=False, encoding='utf_8_sig')               
                        