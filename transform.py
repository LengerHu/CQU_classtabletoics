# coding=utf8
import xlwt
import xlrd
import re

d1='星期一'
d2='星期二'
d3='星期三'
d4='星期四'
d5='星期五'
d6='星期六'
d7='星期日'
d8='星期天'

# xlsx 转移

# 读
readbook = xlrd.open_workbook(r'classtable.xls')
table = readbook.sheet_by_index(0)
# sheet = readbook.sheet_by_name('Sheet0')
nrows = table.nrows #行

#写
writebook = xlwt.Workbook()
sheet = writebook.add_sheet('sheet1')       #在打开的excel中添加一个sheet


sheet.write(0,0,'课程名字')
sheet.write(0,1,'任课教师')
sheet.write(0,2,'上课教室')
sheet.write(0,3,'上课节数1-13')
sheet.write(0,4,'上课时间1-7')
sheet.write(0,5,'开始周')
sheet.write(0,6,'结束周')   


######################################################################
orirownum=2
rownum=1

while 1:
    strvalue=table.cell_value(orirownum,2)

    patternweek=re.compile('(.+)周') 
    weekresult=patternweek.findall(strvalue)
    weekresultlist=weekresult[0].split(',')
    weekloop=1
    weekresultlen=len(weekresultlist)
    while 1:


        #classdaynum 上课节数
        pattern2=re.compile('星期(.+)节')     
        result2=pattern2.findall(strvalue)
        classdaynum=re.findall(r'\d+',str(result2))
        listnum=[]
        i=int(classdaynum[0])
        while 1:
            listnum.append(i)
            if i==int(classdaynum[1]):
                break
            i+=1
        listnum = list(map(str, listnum))
        s = ','.join(listnum)



        #classweeknum 上课时间
        if d3 in strvalue:
            resultofday=3
        if d1 in strvalue:
            resultofday=1
        if d2 in strvalue:
            resultofday=2
        if d4 in strvalue:
            resultofday=4
        if d5 in strvalue:
            resultofday=5
        if d6 in strvalue:
            resultofday=6
        if d7 in strvalue:
            resultofday=7
        if d8 in strvalue:
            resultofday=7

        # 开始周 结束周

        classweeklist=re.findall(r'\d+',str(weekresultlist[weekloop-1]))

        weekbeginnum=int(classweeklist[0])

        if len(classweeklist)==1:
            weekendnum=int(classweeklist[0])
        else:
            weekendnum=int(classweeklist[1])


        sheet.write(rownum,0,table.cell_value(orirownum,0))    #name
        sheet.write(rownum,1,table.cell_value(orirownum,4))    #teacher
        sheet.write(rownum,2,table.cell_value(orirownum,3))    #classroom

        sheet.write(rownum,3,str(s))        #class day num
        sheet.write(rownum,4,resultofday)   #class weekday
        sheet.write(rownum,5,weekbeginnum)  #class week start
        sheet.write(rownum,6,weekendnum)    #class week end

        if weekloop==weekresultlen:
            orirownum+=1
            rownum+=1
            break
        weekloop+=1
        rownum+=1
    if orirownum==nrows:
        break

writebook.save('answer.xls')
print('Then run classtoics.py')








