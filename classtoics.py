from icalendar import Calendar, Event, Alarm
from datetime import datetime, timedelta, date, time
from uuid import uuid1
import pytz
import xlwings

# 读取课表文件
app = xlwings.App(visible=False, add_book=False)
table = app.books.open('answer.xls')
sheet = table.sheets[0]
classnum = sheet.used_range.last_cell.row-1
classinfo = []
for num in range(classnum):
    i = str(num+2)
    singleclassinfo = {}
    singleclassinfo['classname'] = sheet.range('A'+i).value
    singleclassinfo['teacher'] = sheet.range('B'+i).value
    singleclassinfo['location'] = sheet.range('C'+i).value
    try:
        singleclassinfo['classnum'] = sheet.range('D'+i).value.split(',')
    except:
        singleclassinfo['classnum'] = [sheet.range('D'+i).value]
    try:
        singleclassinfo['classday'] = sheet.range('E'+i).value.split(',')
    except:
        singleclassinfo['classday'] = [sheet.range('E'+i).value]
    singleclassinfo['weekstart'] = sheet.range('F'+i).value
    singleclassinfo['weekend'] = sheet.range('G'+i).value
    classinfo.append(singleclassinfo)
table.close()
app.quit()

start_mondey = date(2021, 8, 30)        # 这里每学期要改
term_startyear = start_mondey.year
term_startweek = start_mondey.isocalendar()[1]

#timetable = {'morning':[8,0],'afternoon':[9,0],'evenning':[19,0]}
timetable = {'morning1_2': time(8, 30, 0, 0, tzinfo=pytz.timezone('Asia/Shanghai')),
             'morning3_4': time(10, 30, 0, 0, tzinfo=pytz.timezone('Asia/Shanghai')),
             'afternoon5_6': time(13, 30, 0, 0, tzinfo=pytz.timezone('Asia/Shanghai')),
             'afternoon7_9': time(15, 20, 0, 0,tzinfo=pytz.timezone('Asia/Shanghai')),
             'evenning10_13': time(19, 0, 0, 0,tzinfo=pytz.timezone('Asia/Shanghai'))}

# 创建日历
MyCalender = Calendar()
# 添加属性
MyCalender.add('X-WR-CALNAME', '课程表')
MyCalender.add('prodid', '-//My calendar//luan//CN')
MyCalender.add('version', '2.0')
MyCalender.add('METHOD', 'PUBLISH')
MyCalender.add('CALSCALE', 'GREGORIAN')  # 历法：公历
# MyCalender.add('X-WR-TIMEZONE', 'Asia/Shanghai')  # 通用扩展属性，表示时区
# 循环添加课程
for info in classinfo:
    for day in info['classday']:
        for num in info['classnum']:
            num = int(num)
            event_uuid = str(uuid1())+'@luan'
            event_date = datetime.fromisocalendar(
                term_startyear, term_startweek+int(info['weekstart'])-1, int(day))
            if num <= 2:
                event_start_time = datetime.combine(
                    event_date, timetable['morning1_2']) + timedelta(minutes=(int(num)-1)*55)
            elif num <= 4:
                event_start_time = datetime.combine(
                    event_date, timetable['morning3_4']) + timedelta(minutes=(int(num)-3)*55)
            elif num <= 6:
                event_start_time = datetime.combine(
                    event_date, timetable['afternoon5_6']) + timedelta(minutes=(int(num)-5)*55)
            elif num <= 8:
                event_start_time = datetime.combine(
                    event_date, timetable['afternoon7_9']) + timedelta(minutes=(int(num)-7)*55)
            elif num ==9:
                event_start_time = datetime.combine(
                    event_date, timetable['afternoon7_9']) + timedelta(minutes=(int(num)-7)*60)
            else:
                event_start_time = datetime.combine(
                    event_date, timetable['evenning10_13']) + timedelta(minutes=(int(num)-10)*55)
            event_end_time = event_start_time + timedelta(minutes=45)
            # 添加事件
            event = Event()
            event.add('uid', event_uuid)
            event.add('summary', info['classname'])
            event.add('dtstart', event_start_time)
            event.add('dtend', event_end_time)
            # event.add('dtstamp', datetime.now())
            event.add('location', info['location'])
            event.add('description', '授课老师：'+info['teacher'])
            event.add('rrule', {'freq': 'weekly',
                                'interval': 1,
                                'count': int(info['weekend'])-int(info['weekstart'])+1})
            # 编辑提醒
            if num in [1,2,3,4,5,6,7,8,9,10,11,12]:
                alarm = Alarm()
                alarm.add('action', 'DISPLAY')
                alarm.add('description', 'Reminder')
                alarm.add('trigger;related=start', '-PT5M')
                event.add_component(alarm)
            # elif num in [2, 4, 8, 9]:
            #     alarm = Alarm()
            #     alarm.add('action', 'DISPLAY')
            #     alarm.add('description', 'Reminder')
            #     alarm.add('trigger;related=start', '-PT2M')
            #     event.add_component(alarm)
            # 将事件添加到日历中
            MyCalender.add_component(event)

# 写入文件
with open('课程表.ics', 'wb') as file:
    file.write(MyCalender.to_ical().replace(b'\r\n', b'\n').strip())
    print('导出完毕！')