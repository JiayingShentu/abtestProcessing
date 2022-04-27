import string
from turtle import begin_fill, position

#实验组别对照的字母
group=['A','B','C','D','E','F','G','H']

#行对照的字段
title=['时间-天','AB','弹幕模块UV','弹幕开启UV','弹幕开启率','弹幕默认开启UV','弹幕默认开启率',
'弹幕手动开启UV','弹幕手动开启率','弹幕默认关闭UV','弹幕默认关闭率','弹幕手动关闭UV','弹幕手动关闭率',
'精简模式UV','精简模式开启率','默认精简模式UV','默认精简开启率','手动精简模式UV','手动精简开启率',
'弹幕发送UV','弹幕发送PV','弹幕发送量/弹幕模块UV','人均弹幕发送量','弹幕点赞UV','弹幕点赞PV','弹幕点赞量/弹幕模块UV',
'人均弹幕点赞量','弹幕点踩UV','弹幕点踩PV','弹幕点踩量/弹幕模块UV','人均点踩量','弹幕点赞比例','弹幕发布比例']

type1=['弹幕模块UV','弹幕开启UV','弹幕默认开启UV','弹幕手动开启UV','弹幕默认关闭UV','弹幕手动关闭UV',
    '精简模式UV','默认精简模式UV','手动精简模式UV','弹幕发送UV','弹幕发送PV','弹幕点赞UV','弹幕点赞PV',
    '弹幕点踩UV','弹幕点踩PV']
type2=['弹幕开启率','弹幕默认开启率','弹幕手动开启率','弹幕默认关闭率','弹幕手动关闭率','精简模式开启率',
'默认精简开启率','手动精简开启率','弹幕发送量/弹幕模块UV','人均弹幕发送量','弹幕点赞量/弹幕模块UV',
'人均弹幕点赞量','弹幕点踩量/弹幕模块UV','人均点踩量','弹幕点赞比例','弹幕发布比例']
#该函数确定取数的时间段&天数
def getBasicData(sheet):
    dateNum=1   # 取了dateNum天的数
    testNum=0   # 实验共testNum组
    date=[]     # 取数的时间段
    test=[]     # 组别名称
    colA=sheet[title[0]]
    colB=sheet[title[1]]
    tmp=''  
    #testNum计算 
    for item in colA:
        if tmp==item or tmp=='':
            testNum=testNum+1
        else:
            break
        tmp=item
    #test确定
    tmp=colB[1][:-1]
    for i in range(testNum):
        test.append(tmp+group[i])   
    #dateNum计算
    dateNum=int((len(sheet))/testNum)
    #date确定
    for i in range(dateNum):
        date.append(colA[testNum*i])
    return test,date
       
#该函数写入第一列（日期列）
def printCol_1(sheet,date):
    sheet['A2']=title[0]
    for i in range(len(date)):
        sheet['A'+str(i+3)]=date[i]

#该函数写入后续数据
def printTable(sheet,df,date,test,colName):
    if colName=='弹幕模块UV':
        position=sheet.max_column
    else:
        position=sheet.max_column+1
    #不需要分别与AB组对比的情况
    if colName in type1:
        easyMode(sheet,df,date,test,colName,position)
    #需要分别与AB组对比的情况
    if colName in type2:
        hardMode(sheet,df,date,test,colName,position)
        


def easyMode(sheet,df,date,test,colName,position):
    sheet.cell(row=1,column=position+1).value=colName
    for i in range(len(test)):
        sheet.cell(row=2,column=position+i+1).value=test[i]
        for j in range(len(date)):
            sheet.cell(row=j+3,column=position+i+1).value=df.at[(date[j],test[i]),colName]

def hardMode(sheet,df,date,test,colName,position):
    sheet.cell(row=1,column=position+1).value=colName
    #写入对照组A组
    sheet.cell(row=2,column=position+1).value=test[0]
    for j in range(len(date)):
        aCellValue=df.at[(date[j],test[0]),colName]
        sheet.cell(row=j+3,column=position+1).value=aCellValue
    #写入实验组&对比数据
    for i in range(len(test)-2):
        sheet.cell(row=2,column=position+2+i).value=test[i+2]
        for j in range(len(date)):
            cellValue=df.at[(date[j],test[i+2]),colName]
            sheet.cell(row=j+3,column=position+2+i).value=cellValue
            compare=dataCompare(aCellValue,cellValue)
            sheet.cell(row=len(date)+3+8+j,column=position+2+i).value=compare
    #重置position位置
    position=position+len(test)
    #写入对照组B组
    sheet.cell(row=2,column=position+1).value=test[1]
    for j in range(len(date)):
        bCellValue=df.at[(date[j],test[1]),colName]
        sheet.cell(row=j+3,column=position+1).value=bCellValue
    #写入实验组
    for i in range(len(test)-2):
        sheet.cell(row=2,column=position+2+i).value=test[i+2]
        for j in range(len(date)):
            cellValue=df.at[(date[j],test[i+2]),colName]
            sheet.cell(row=j+3,column=position+2+i).value=cellValue
            compare=dataCompare(bCellValue,cellValue)
            sheet.cell(row=len(date)+3+8+j,column=position+2+i).value=compare


def dataCompare(str1,str2):
    if type(str1)==str:
        if str1[-1]=='%':
            result=float(str2.strip('%'))-float(str1.strip('%'))
            print (str2,'-',str1,'=',str(round(result,2))+'%')
            return str(round(result,2))+'%'
    result=float(str2)-float(str1)
    return str(round(result))

        


    
    


            

