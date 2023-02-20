import datetime
from icalendar import Calendar, Event
import openpyxl

print("程序初始化中……")
#输入学号
userid=input('请输入学号：\n')
filepath="./学生个人课表_"+userid+".xlsx"

#打开xlsx文件
WorkBook=openpyxl.load_workbook(filepath)
print(WorkBook.sheetnames)