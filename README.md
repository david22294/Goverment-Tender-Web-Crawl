# Goverment-Tender-Web-Crawl
Goverment Tender Web Crawl
  
NOTE: openpyxl == 3.0.0  
  

Relaese v2.x
Use to craw tender form https://web.pcc.gov.tw/prkms/tender/common/basic/indexTenderBasic.
Only craw tender before "各級學校"  

How to use:  
  1. Click dist/main.exe to run program.  
  2. Program will collect tender since a week ago as default.  
  3. Output will save as excel.  
![image](https://github.com/david22294/Goverment-Tender-Web-Crawl/blob/main/example/image/ExcelOutput.PNG)

[Feature] Filter
Can add exception in "機關名稱" "招標方式 "採購性質" in ExceptionList.josn.
Default is empty.
How to add exception list:
Ex: "採購性質":["工程類"] => "工程類" will not save into output excel.

-------------------------------------------------------------------------------------------------------------------------
Relaese v1.x
Use to craw tender form https://web.pcc.gov.tw/tps/main/pss/pblm/tender/basic/search/mainListCommon.jsp?searchType=basic.  
Only craw tender before "各級學校"  

How to use:  
  1. Click dist/main.exe to run program.  
  2. Program will collect tender since a week ago as default.  
  3. Output will save as excel.  
![image](https://github.com/david22294/Goverment-Tender-Web-Crawl/blob/main/example/image/ExcelOutput.PNG)

[Feature] Filter
Can add exception in "機關名稱" "招標方式 "採購性質" in ExceptionList.josn.
Default is empty.
How to add exception list:
Ex: "採購性質":["工程類"] => "工程類" will not save into output excel.
