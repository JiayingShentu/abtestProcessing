from numpy import true_divide
from openpyxl import load_workbook
from openpyxl import Workbook
import pandas as pd
from function import getBasicData,printCol_1,printTable

#读取abtest原表格
originData=pd.read_excel(io='./origin.xlsx')
oD=originData.set_index(['时间-天', 'AB'])

#创建abtest新表格，用于分析数据
newWorkBook=Workbook()
ws=newWorkBook.active
ws.title="数据结果"
#获取确定取数的时间段&实验组别
test,date=getBasicData(originData)

printCol_1(ws,date)     #写入第一列（日期列）
printTable(ws,oD,date,test,'弹幕模块UV')  #写入'弹幕模块UV'结果



newWorkBook.save("example.xlsx")