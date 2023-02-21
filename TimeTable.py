import datetime
import re
from icalendar import Calendar, Event, Alarm
import openpyxl
import time

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
#userid=input('请输入学号，确保和xlsx文件名中的学号一致：')
userid='2021212702'
WorkBook=openpyxl.load_workbook(filename=f"./学生个人课表_{userid}.xlsx")
Sheet=WorkBook.active

print("您的信息为：")
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
time.sleep(3)


'''
#输入学期的第一周的周一日期
Start=datetime.datetime.strptime(input("输入学期的第一周的周一日期，以YYYY-MM-DD格式\n"), '%Y-%m-%d').date()
while Start.isoweekday() != 1:
    Start=datetime.datetime.strptime(input("日期并非周一！请以YYYY-MM-DD格式输入\n"), '%Y-%m-%d').date()
print("正在处理，请稍等")
'''
StartDay=datetime.date(2023, 2, 20)

#制作
MyCalendar = Calendar()
MyCalendar.add('X-WR-CALNAME', SchoolYear)
MyCalendar.add('X-APPLE-CALENDAR-COLOR', '#E1FFFF')
MyCalendar.add('X-WR-TIMEZONE', 'Asia/Shanghai')
MyCalendar.add('VERSION', '2.0')
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
            if i/5==0:
                MyEvent=Event()
                MyEvent.add('UID', f'BUPTCalendar@{StudentName}&{datetime.datetime.now().timestamp()}')
                MyEvent.add('SUMMARY', Course)
                #MyEvent.add('DTSTART', StartTime[row-4])
                #MyEvent.add('DTEND', EndTime[i-4])

                MyAlarm=Alarm()
                MyAlarm.add('TRIGGER', "-PT10M")
                MyAlarm.add('ACTION', "DISPLAY")
                MyAlarm.add('DESCRIPTION', Course)
                MyEvent.add_component(MyAlarm)
                MyCalendar.add_component(MyEvent)
                del MyAlarm
                del MyEvent
                print(f"\n-----------周{column-1}第{row-3}节课程导入完成-----------", end="")
        del CellBR


'''
    try:
        with open('TimeTable.ics', 'wb') as file:
            file.write(cal.to_ical())
            print('[Success]')
    except Exception:
        print("生成文件失败，请重试")
'''    
