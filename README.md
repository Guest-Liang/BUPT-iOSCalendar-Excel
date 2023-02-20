# BUPT-iOSCalendar-Excel
使用从北邮教务系统中导出的Excel版课程表创建ics格式的日历，导入到Apple设备中

## 使用方法
### 第一步
需要`python`环境，然后在`cmd`或者`PowerShell`运行以下代码，安装需要的库：
```python3
pip install icalendar
```
### 第二步
从你的北邮教务里下载excel版个人课程表，文件名为“学生个人课表_{你的学号}.xls”
![Alt text](https://github.com/Guest-Liang/BUPT-iOSCalendar-Excel/blob/main/ScrennShots/GetExcelFile.png)

### 第三步
下载`py文件`，将`Excel`文件和`py文件`放在同一个目录下，运行`py文件`

### 第四步
得到的`ics文件`导入Apple设备中即可使用
