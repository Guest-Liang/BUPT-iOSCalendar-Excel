import datetime
import re
from icalendar import Calendar, Event
import openpyxl

print("程序初始化中……")
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


#获取学号，打开xlsx文件
#userid=input('请输入学号：')
userid='2021212702'
filepath="./学生个人课表_"+userid+".xlsx"
WorkBook=openpyxl.load_workbook(filepath)
Sheet=WorkBook.active

print("将要处理的课程表所属人：", end="")
Name=Sheet['A1'].value.replace("北京邮电大学 ","").replace(" 学生个人课表","")
print(Name)

print("学年：",end="")
SchoolYear=Sheet['A2'].value[5:16]
print(SchoolYear)

print("班级：",end="")
TheClass=Sheet['A2'].value[27:37]
print(TheClass)

print("专业：",end="")
Major=Sheet['A2'].value[48:55]
print(Major)

print("学院：",end="")
Academy=Sheet['A2'].value[66:70]
print(Academy)


#处理信息
#找到字符串中某个字的索引
def GetElementIndex(string, char):
    return [idx.start() for idx in re.finditer(char, string)]

#输入学期的第一周的周一日期
Start=datetime.datetime.strptime(input("输入学期的第一周的周一日期，以YYYY-MM-DD格式\n"), '%Y-%m-%d').date()
while Start.isoweekday() != 1:
    Start=datetime.datetime.strptime(input("日期并非周一！请以YYYY-MM-DD格式输入\n"), '%Y-%m-%d').date()

    


#制作ics文件
def MakeicsFile():
    cal = Calendar()
    cal.add('X-WR-CALNAME', SchoolYear)
    cal.add('X-APPLE-CALENDAR-COLOR', '#E1FFFF')
    cal.add('X-WR-TIMEZONE', 'Asia/Beijing')
    cal.add('VERSION', '2.0')