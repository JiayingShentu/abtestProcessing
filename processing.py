from openpyxl import load_workbook
from openpyxl import Workbook
from function import getBasicData,printCol_1,printTable

#读取abtest原表格
originWB=load_workbook(filename='origin.xlsx')
sheet1=originWB['Sheet1']
#创建abtest新表格，用于分析数据
doneWB=Workbook()
ws=doneWB.active
ws.title="数据结果"
#获取确定取数的时间段&实验组别
testNum,test,dateNum,date=getBasicData(sheet1)

printCol_1(ws,date)     #写入第一列（日期列）
printTable(ws,sheet1,date,test)  #写入'弹幕模块'UV结果

doneWB.save("example.xlsx")
