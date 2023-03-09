# BUPT-iOSCalendar-Excel
使用从北邮教务系统中导出的Excel版课程表创建ics格式的日历，导入到Apple设备中 【维护中，有问题提issue，看到会解决】

## 使用方法
### 第一步
需要`高于python3.10`的`python`环境，并且配置好环境变量等，确保`Powershell`中输入`python`能出现版本号并进入python环境。如果已经配置好，在`cmd`或者`PowerShell`运行以下代码，安装需要的库：
```python3
pip install icalendar
pip install openpyxl
```   
连接不上可使用清华源：
```python3
pip install icalendar -i https://pypi.tuna.tsinghua.edu.cn/simple
pip install openpyxl -i https://pypi.tuna.tsinghua.edu.cn/simple
```

### 第二步
从你的北邮教务里下载Excel版个人课程表，文件名为“学生个人课表_{你的学号}.xls”  
在Excel中将其另存为为xlsx格式，保存后你的文件名应该是“学生个人课表_{你的学号}.xlsx”   
<img src="https://github.com/Guest-Liang/BUPT-iOSCalendar-Excel/blob/main/ScrennShots/GetExcelFile.png" width="500px">

### 第三步
在GitHub页面右边的`release`下载`TimeTable.py`文件，将`学生个人课表_{你的学号}.xlsx`文件和`TimeTable.py`文件放在同一个目录下，在当前目录空白处右键“在终端中打开”，或者打开`Powershell`，进入管理员模式，执行
```python3
python TimeTable.py
```
<img src="https://github.com/Guest-Liang/BUPT-iOSCalendar-Excel/blob/main/ScrennShots/ExecuteTheCommand.png" width="500px">   
   
按提示输入**你的学号**，以及本学期**第一周周一的日期**，等待运行  

看到最后的Success就说明成功了，ics文件生成在当前目录下  
<img src="https://github.com/Guest-Liang/BUPT-iOSCalendar-Excel/blob/main/ScrennShots/Success.png" width="500px">   
如果失败了请在issue中提出，并附上你的xlsx文件，以及所使用的python版本号等一切必要的信息   


### 第四步
得到的带有名字的`ics文件`导入Apple设备中即可使用。  
推荐添加到一个新的日历：以学年命名或者学习，这样万一添加错误还可以通过删除整个日历来重新添加，不需要一个个手动删除   
**建议在日历中新建好新的日历再打开ics文件添加**   
确保在添加到日历前全部检查一遍，不然需要重新添加   
有问题千万不要导入！  


# 有问题去提issue（等待中）
# 目前bug：  
iOS & iPadOS不能识别私有属性中的颜色，导致`X-APPLE-CALENDAR-COLOR`这一项参数无效   


## 咕咕咕中：
1、正在考虑实现利用学号、web登录、教务密码直接从教务系统中获取课程表（开发中……但是咕咕咕）   
2、One more thing……  
