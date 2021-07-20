# 重庆大学课表 Excel 文件转 ics
## 简介
此项目可以将重庆大学教务管理系统下载的课表 excel 转为 ics 导入到支持 ics 导入的日历中，导出的 ics 文件经测试可以在 iOS、MacOS、Windows、以及基于 Android 的 RealmeUI 的内置日历中使用
## 使用方法
1.进入[重庆大学教务管理系统](http://my.cqu.edu.cn/enroll/CourseStuSelectionList)，登录后点击查看课表，右上方点击 Excel 下载按钮

2.打开下载的课表，另存为 Excel 97-2003 工作簿 （后缀名为 .xls），修改文件名为 classtable

3.下载本项目的两个py文件，将另存为得到的课表文件 classtable.xls 放在本项目同一目录（文件夹)中

4.先运行 transform.py ，再运行 classtoics.py

5.ics在该目录下就生成辣！可以自行导入到日历中
