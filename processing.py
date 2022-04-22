from openpyxl import load_workbook
from openpyxl import Workbook

#读取abtest原表格
originWB=load_workbook(filename='origin.xlsx')
sheet1=originWB['Sheet1']
#创建abtest新表格，用于分析数据
doneWB=Workbook()
ws=doneWB.active
ws.title="数据结果"

#groupNum存储实验+对照的组数
#groupNum=int(input('你的AB实验数据有几组:'))
group=['A','B','C','D','E','F','G','H']
for item in sheet1.rows:
        print(item)  # 按照行输出
        print(item[0].value)  # 输出单元格的值
        item[0].value = 1
# 输出修改后的值，注意不保存文件，数据不会存储
for item in ws.rows:
    print(item[0].value)  # 输出单元格的值

#doneWB.save("example.xlsx")


