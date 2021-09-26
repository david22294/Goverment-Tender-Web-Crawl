import json
import requests
import htmls
from bs4 import BeautifulSoup
from lxml import etree
import time
import pandas
import datetime 


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

def createPostData(orgId, StartTime, EndTime):
    postForm = {'method':'search',
     'searchMethod':'true',
     'orgId':orgId,
     'hid_1':'1',
     'tenderType':'tenderDeclaration',
     'tenderWay': '1,2,3,4,5,6,7,10,12',
     'tenderDateRadio':'on',
     'tenderStartDateStr': StartTime,
     'tenderEndDateStr': EndTime,
     'tenderStartDate': StartTime,
     'tenderEndDate': EndTime,
     'isSpdt':'N',
     'btnQuery':'查詢'
    }
    return postForm

def getColumnInf(column, var):
    return {
        'department' : column[1].text,
        'tenderNum' : column[2].text.split("\r\n")[1].replace('\t',''),
        'tenderName' : column[2].find("a").get("title"),
        'tenderMethod' : column[4].text.replace('\t','').replace('\n','').replace('\r',''),
        'tenderType' : column[5].text.replace('\t','').replace('\n','').replace('\r',''),
        'annouceDate' : column[6].text.replace('\t','').replace('\n','').replace('\r',''),
        'deadline' : column[7].text.replace('\t','').replace('\n','').replace('\r',''),
        'money' : column[8].text.replace('\t','').replace('\n','').replace('\r',''),
        'trendHref' : "https://web.pcc.gov.tw/tps{i}\n".format(i=column[2].find("a").get("href")[2:]),
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

basePath = "https://web.pcc.gov.tw/tps/main/pss/pblm/tender/basic/search/"
buyPostPath = "https://web.pcc.gov.tw/tps/pss/tender.do?searchMode=common&searchType=basic"

if __name__ == '__main__':
    
    dayRange = input("How long do you want to colloect?(day) ")
    
    exceptionList = getException()
    
    currentTime = datetime.datetime.now()
    weekAgoTime = currentTime - datetime.timedelta(days = int(dayRange))
    StartTime = "{Y}/{M}/{D}".format(Y=weekAgoTime.year-1911, M=weekAgoTime.month, D=weekAgoTime.day)
    StartTime = "{Y}/{MD}".format(Y=weekAgoTime.year-1911, MD=weekAgoTime.strftime("%m/%d"))
    EndTime = "{Y}/{MD}".format(Y=currentTime.year-1911, MD=currentTime.strftime("%m/%d"))
    
    print("Start Time: "+StartTime)
    print("End Time: "+EndTime)
    
    root = requests.get("https://web.pcc.gov.tw/tps/main/pss/pblm/tender/basic/search/mainListCommon.jsp?searchType=basic")
    rootContent = root.content.decode()
    rootHtmltml = etree.HTML(rootContent)
    departments = rootHtmltml.xpath("//p/a[@style='color: blue;']")
    print(len(departments))
#     x = 0
    for department in departments:
#         x = x + 1
#         if (x == 5): break;
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
                postData = createPostData(orgId, StartTime, EndTime)
                buy = requests.post(buyPostPath, data=postData)
                time.sleep(2) # 

                buySoup = BeautifulSoup(buy.text, "html.parser")
                items = buySoup.find_all("tr", onmouseover="overcss(this);")
                column = None
                if orgId == None:
                    print('\t'+departmentName+' '+str(orgId)+' => None')
                elif len(items)<1:
                    print('\t'+departmentName+' '+str(orgId)+' => None')
                else:
                    print('\t'+departmentName+' '+str(orgId)+' => '+str(len(items)))
                    
                    for item in items:
                        
                        isPass = False
                        infs = []
                        
                        # get tend inf
                        column = item.find_all("td")
                        for name in columnName:
                            if not isPass:
                                var = getColumnInf(column, name)
                                infs.append(var)
                                if name in exceptionList.keys():
                                    if var in exceptionList[name]:
                                        isPass = True
                        
                        if not isPass:
                            for i, name in enumerate(columnName):
                                outputColumn[chineseColumnName[name]].append(infs[i])
    df = pandas.DataFrame(outputColumn)
    df.to_excel("國家機關採購案.xlsx",index=False, encoding='utf_8_sig')               
                        