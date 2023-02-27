# BUPT-iOSCalendar-Excel
使用从北邮教务系统中导出的Excel版课程表创建ics格式的日历，导入到Apple设备中【持续更新bug中】

## 使用方法
### 第一步
需要`>python3.10`环境，并且配置好环境变量等，确保`Powershell`中输入`python`能出现版本号并进入python环境。如果已经配置好，在`cmd`或者`PowerShell`运行以下代码，安装需要的库：
```python3
pip install icalendar
pip install openpyxl
```
### 第二步
从你的北邮教务里下载Excel版个人课程表，文件名为“学生个人课表_{你的学号}.xls”。  
在Excel中将其另存为为xlsx格式，保存后你的文件名应该是“学生个人课表_{你的学号}.xlsx”
<img src="https://github.com/Guest-Liang/BUPT-iOSCalendar-Excel/blob/main/ScrennShots/GetExcelFile.png" width="500px">

### 第三步
下载`TimeTable.py`文件，将`学生个人课表_{你的学号}.xlsx`文件和`TimeTable.py`文件放在同一个目录下，在空白处右键“在终端中打开”，或者打开`Powershell`，进入管理员模式，执行
```python3
python TimeTable.py
```
<img src="https://github.com/Guest-Liang/BUPT-iOSCalendar-Excel/blob/main/ScrennShots/ExecuteTheCommand.png" width="500px">  
按提示输入你的学号，以及本学期第一周周一的日期  

看到最后的Success就说明成功了，ics文件生成在当前目录下  
<img src="https://github.com/Guest-Liang/BUPT-iOSCalendar-Excel/blob/main/ScrennShots/Success.png" width="500px">


### 第四步
得到的`ics文件`导入Apple设备中即可使用。  
确保在添加到日历前全部检查一遍，不然手动删除会很麻烦！  
确保在添加到日历前全部检查一遍，不然手动删除会很麻烦！  
确保在添加到日历前全部检查一遍，不然手动删除会很麻烦！  
有问题千万不要导入！否则删除非常麻烦！  
有问题千万不要导入！否则删除非常麻烦！  
有问题千万不要导入！否则删除非常麻烦！  

# 有问题去提issue
# 目前bug：  
PC端Outlook正常识别全部课程  
测试用iPhone 7 Plus正常识别(iOS15.7.3)  
测试用iPad Pro (11-inch) (Gen3)正常识别(iPadOS16.3.1)


## 咕咕咕中：
1、将相邻两节课/三节课甚至更多的课合并为一个事件  
2、正在考虑实现利用学号、web登录、教务密码直接从教务系统中获取课程表（咕咕咕）  
3、One more thing…  
