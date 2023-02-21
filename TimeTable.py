import datetime
import re
from icalendar import Calendar, Event
import openpyxl

print("程序初始化中………")
#定义课程开始时间
StartTime=[datetime.time(8, 0, 0), datetime.time(8, 50, 0), datetime.time(9, 50, 0), 
           datetime.time(10, 40, 0), datetime.time(11, 30, 0), datetime.time(13, 00, 0), 
           datetime.time(13, 50, 0), datetime.time(14, 45, 0), datetime.time(15, 40, 0), 
           datetime.time(16, 35, 0), datetime.time(17, 25, 0), datetime.time(18, 30, 0), 
           datetime.time(19, 20, 0), datetime.time(20, 10, 0)]
#定义课程结束时间
EndTime=[datetime.time(8, 45, 0), datetime.time(9, 35, 0), datetime.time(10, 35, 0), 
         datetime.time(11, 25, 0), datetime.time(12, 15, 0), datetime.time(13, 45, 0), 
         datetime.time(14, 35, 0), datetime.time(15, 30, 0), datetime.time(16, 25, 0), 
         datetime.time(17, 20, 0), datetime.time(18, 10, 0), datetime.time(19, 15, 0), 
         datetime.time(20, 5, 0), datetime.time(20, 55, 0)]
#找到字符串中某个关键值的所有索引
def GetElementIndex(char, string):
    return [idx.start() for idx in re.finditer(char, string)]



#获取学号，打开xlsx文件
#userid=input('请输入学号：')
userid='2021212702'
WorkBook=openpyxl.load_workbook(filename=f"./学生个人课表_{userid}.xlsx")
Sheet=WorkBook.active

print("-------------------------")
print("课程表所属人：", end="")
print(Sheet['A1'].value.replace("北京邮电大学 ","").replace(" 学生个人课表",""))
StudentName=Sheet['A1'].value.replace("北京邮电大学 ","").replace(" 学生个人课表","")

print("学年：",end="")
print(Sheet['A2'].value[5:16])
SchoolYear=Sheet['A2'].value[5:16]

print("班级：",end="")
print(Sheet['A2'].value[27:37])

print("专业：",end="")
print(Sheet['A2'].value[48:55])

print("学院：",end="")
print(Sheet['A2'].value[66:70])
print("-------------------------")



'''
#输入学期的第一周的周一日期
Start=datetime.datetime.strptime(input("输入学期的第一周的周一日期，以YYYY-MM-DD格式\n"), '%Y-%m-%d').date()
while Start.isoweekday() != 1:
    Start=datetime.datetime.strptime(input("日期并非周一！请以YYYY-MM-DD格式输入\n"), '%Y-%m-%d').date()
print("正在处理，请稍等")
'''
StartDay=datetime.date(2023, 2, 20)

#制作
Cal = Calendar()
Cal.add('X-WR-CALNAME', SchoolYear)
Cal.add('X-APPLE-CALENDAR-COLOR', '#E1FFFF')
Cal.add('X-WR-TIMEZONE', 'Asia/Shanghai')
Cal.add('VERSION', '2.0')
for row in range(4, 18):
    for column in range(2, 9):
        CellBR=GetElementIndex("\n", Sheet.cell(row, column).value)
        for i in range(len(CellBR)-1):
            match i%5:
                case 0:
                    Course=Sheet.cell(row, column).value[CellBR[i]:CellBR[i+1]]
                case 1:
                    TeacherName=Sheet.cell(row, column).value[CellBR[i]:CellBR[i+1]]
                case 2:
                    ClassWeeks=Sheet.cell(row, column).value[CellBR[i]:CellBR[i+1]]
                case 3:
                    Classroom=Sheet.cell(row, column).value[CellBR[i]:CellBR[i+1]]
                case 4:
                    LessonNum=Sheet.cell(row, column).value[CellBR[i]:CellBR[i+1]]
            #print(Sheet.cell(row, column).value[CellBR[i]:CellBR[i+1]], end="")
            if i==len(CellBR)-2:
                LessonNum=Sheet.cell(row, column).value[CellBR[-1]:]
                #print(f"{Sheet.cell(row, column).value[CellBR[-1]:]}", end="")

        event=Event()
        event.add('uid', f'BUPT@{StudentName}&{datetime.datetime.now().timestamp()}')
        # event.add('dtstart', dtstart)
        # event.add('dtend', dtend)
        # event.add('summary', lesson['name'])


        # print(f"\n-----------周{column-1}第{row-3}节课程拆分-----------", end="")
        del CellBR








def AddEvent(Cal, SUMMARY, DTSTART, DTEND):
    """
    向Calendar日历对象添加事件的方法
    :param cal: calender日历实例
    :param SUMMARY: 事件名
    :param DTSTART: 事件开始时间
    :param DTEND: 时间结束时间
    :return:
    """
    time_format = "TZID=Asia/Shanghai:{date.year}{date.month:0>2d}{date.day:0>2d}T{date.hour:0>2d}{date.minute:0>2d}00"
    dt_start = time_format.format(date=DTSTART)
    dt_end = time_format.format(date=DTEND)
    CreateTime = datetime.datetime.today().strftime("%Y%m%dT%H%M%SZ")
    Cal.add_event(
        SUMMARY=SUMMARY,
        DTSTART=dt_start,
        DTEND=dt_end,
        DTSTAMP=CreateTime,
        SEQUENCE="0",
        CREATED=CreateTime,
        LAST_MODIFIED=CreateTime,
        STATUS="CONFIRMED",
    )




#制作ics文件
MyCalendar = Calendar(CalendarName=f"{SchoolYear}课程表")
AddEvent(MyCalendar,
        SUMMARY="测试",
        DTSTART=datetime.datetime(year=2023,month=2,day=20,hour=21,minute=20,second=00),
        DTEND=datetime.datetime(year=2023,month=2,day=20,hour=21,minute=30,second=00),)







'''
    try:
        with open('TimeTable.ics', 'wb') as file:
            file.write(cal.to_ical())
            print('[Success]')
    except Exception:
        print("生成文件失败，请重试")
'''    
