import requests
import htmls
from bs4 import BeautifulSoup
from lxml import etree
import time
import pandas
import datetime


def createPostData(orgId, StartTime, EndTime):
    postForm = {'method': 'search',
                'searchMethod': 'true',
                'orgId': orgId,
                'hid_1': '1',
                'tenderType': 'tenderDeclaration',
                'tenderWay': '1,2,3,4,5,6,7,10,12',
                'tenderDateRadio': 'on',
                'tenderStartDateStr': StartTime,
                'tenderEndDateStr': EndTime,
                'tenderStartDate': StartTime,
                'tenderEndDate': EndTime,
                'isSpdt': 'N',
                'btnQuery': '查詢'
                }
    return postForm


basePath = "https://web.pcc.gov.tw/tps/main/pss/pblm/tender/basic/search/"
buyPostPath = "https://web.pcc.gov.tw/tps/pss/tender.do?searchMode=common&searchType=basic"

departmentA = []
tenderNum = []
tenderName = []
tenderMethod = []
tenderType = []
announceDate = []
deadline = []
money = []
trendHref = []

currentTime = datetime.datetime.now()
weekAgoTime = currentTime - datetime.timedelta(days=6)
StartTime = "{Y}/{M}/{D}".format(Y=weekAgoTime.year -
                                 1911, M=weekAgoTime.month, D=weekAgoTime.day)
StartTime = "{Y}/{MD}".format(Y=weekAgoTime.year -
                              1911, MD=weekAgoTime.strftime("%m/%d"))
EndTime = "{Y}/{MD}".format(Y=currentTime.year-1911,
                            MD=currentTime.strftime("%m/%d"))

if __name__ == '__main__':
    root = requests.get(
        "https://web.pcc.gov.tw/tps/main/pss/pblm/tender/basic/search/mainListCommon.jsp?searchType=basic")
    rootContent = root.content.decode()
    rootHtml = etree.HTML(rootContent)
    departments = rootHtml.xpath("//p/a[@style='color: blue;']")
    print(len(departments))
    for department in departments:
        print(department.attrib['title'])
        if department.attrib['title'] == '各級學校':
            print('Stop at 各級學校')
            break
        targetPath = basePath + department.attrib['href']
        targetRoot = requests.get(targetPath)
        targetSoup = BeautifulSoup(targetRoot.text, "html.parser")
        buyList = targetSoup.find(
            "table", align='center', width='100%', border='0').findChildren("tr")
        for index in range(2, len(buyList)-1):
            itemColumn = buyList[index].findChildren('td')
            if len(itemColumn) < 2:
                pass
            else:
                orgId = itemColumn[0].string
                departmentName = itemColumn[1].string
                postData = createPostData(orgId, StartTime, EndTime)
                buy = requests.post(buyPostPath, data=postData)
                buySoup = BeautifulSoup(buy.text, "html.parser")
                items = buySoup.find_all("tr", onmouseover="overcss(this);")
                column = None
                if orgId == None:
                    print('\t'+departmentName+' '+str(orgId)+' => PASS')
                    time.sleep(2)
                    pass
                elif len(items) < 1:
                    print('\t'+departmentName+' '+str(orgId)+' => PASS')
                    time.sleep(1)
                    pass
                else:
                    print('\t'+departmentName+' ' +
                          str(orgId)+' => '+str(len(items)))
                    for item in items:
                        column = item.find_all("td")
                        # print("\t 機關名稱: {i}".format(i=column[1].text))
                        # print("\t 檔案案號: {i}".format(i=column[2].text.split("\r\n")[1].replace('\t','')))
                        # print("\t 標案名稱: {i}".format(i=column[2].find("a").get("title")))
                        # print("\t 招標方式: {i}".format(i=column[4].text.replace('\t','').replace('\n','').replace('\r','')))
                        # print("\t 採購性質: {i}".format(i=column[5].text.replace('\t','').replace('\n','').replace('\r','')))
                        # print("\t 公告日期: {i}".format(i=column[6].text.replace('\t','').replace('\n','').replace('\r','')))
                        # print("\t 截止投標: {i}".format(i=column[7].text.replace('\t','').replace('\n','').replace('\r','')))
                        # print("\t 預算金額: {i}".format(i=column[8].text.replace('\t','').replace('\n','').replace('\r','')))
                        # print("\t 連結: https://web.pcc.gov.tw/tps{i}\n".format(i=column[2].find("a").get("href")[2:]))
                        departmentA.append(column[1].text)
                        tenderNum.append(column[2].text.split(
                            "\r\n")[1].replace('\t', ''))
                        tenderName.append(column[2].find("a").get("title"))
                        tenderMethod.append(column[4].text.replace(
                            '\t', '').replace('\n', '').replace('\r', ''))
                        tenderType.append(column[5].text.replace(
                            '\t', '').replace('\n', '').replace('\r', ''))
                        announceDate.append(column[6].text.replace(
                            '\t', '').replace('\n', '').replace('\r', ''))
                        deadline.append(column[7].text.replace(
                            '\t', '').replace('\n', '').replace('\r', ''))
                        money.append(column[8].text.replace(
                            '\t', '').replace('\n', '').replace('\r', ''))
                        trendHref.append(
                            "https://web.pcc.gov.tw/tps{i}\n".format(i=column[2].find("a").get("href")[2:]))
                        time.sleep(1)

    df = pandas.DataFrame({'機關名稱': departmentA,
                           '檔案案號': tenderNum,
                           '標案名稱': tenderName,
                           '招標方式': tenderMethod,
                           '採購性質': tenderType,
                           '公告日期': announceDate,
                           '截止投標': deadline,
                           '預算金額': money,
                           '連結': trendHref})
    df.to_excel("國家機關採購案.xlsx", index=False, encoding='utf_8_sig')
