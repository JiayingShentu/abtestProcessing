#实验组别对照的字母
group=['A','B','C','D','E','F','G','H']

#行对照的字段
title=['时间-天','AB','弹幕模块UV','弹幕开启UV','弹幕开启率','弹幕默认开启UV','弹幕默认开启率',
'弹幕手动开启UV','弹幕手动开启率','弹幕默认关闭UV','弹幕默认关闭率','弹幕手动关闭UV','弹幕手动关闭率',
'精简模式UV','精简模式开启率','默认精简模式UV','默认精简开启率','手动精简模式UV','手动精简开启率',
'弹幕发送UV','弹幕发送PV','弹幕发送量/弹幕模块UV','人均弹幕发送量','弹幕点赞UV','弹幕点赞PV','弹幕点赞量/弹幕模块UV',
'人均弹幕点赞量','弹幕点踩UV','弹幕点踩PV','弹幕点踩量/弹幕模块UV','人均点踩量','弹幕点赞比例','弹幕发布比例']

#该函数确定取数的时间段&天数
def getBasicData(sheet):
    dateNum=1   # 取了dateNum天的数
    testNum=1   # 实验共testNum组
    date=[]     # 取数的时间段
    test=[]     # 组别名称
    colA=sheet['A']
    colB=sheet['B']
    tmp=''  
    #testNum计算 
    for cell in colA:
        if cell.value==title[0]:
            print('yes')
        if cell.value!=title[0] and tmp!=title[0]:
            if tmp==cell.value:
                testNum=testNum+1
            else:
                break
        tmp=cell.value
    #test确定
    tmp=colB[1].value[:-1]
    for i in range(testNum):
        test.append(tmp+group[i])   
    #dateNum计算
    dateNum=int((sheet.max_row-1)/testNum)
    #date确定
    for i in range(dateNum):
        date.append(colA[1+testNum*i].value)
    return testNum,test,dateNum,date

#该函数写入第一列（日期列）
def printCol_1(sheet,date):
    sheet['A1']=title[0]
    for i in range(len(date)):
        sheet['A'+str(i+2)]=date[i]

#该函数写入后续数据
def printTable(sheet,sheet1,date,test):
    data=[]
    for r in range(1,sheet1.max_row):
        if(sheet1['A'+str(r+1)]==date[0])

