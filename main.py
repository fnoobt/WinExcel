from tkinter import *
from mainpage import * 

#-F 表示生成单个可执行文件,-w 表示去掉控制台窗口,-i 表示可执行文件的图标
#pyinstaller -i logo\EBO.ico -F -w main.py
#pyinstaller -p D:\users\cmcc\AppData\Local\Programs\Python\Python36-32\Lib\site-packages\win32\lib -p D:\PythonProject\pytkinter\winexcel -i logo\EBO.ico -w -F main.py
#Excel batch operation,EBO,Excel批量操作助手
root = Tk()
root.title('Excel Batch Operation Tool')
MainPage(root)
root.mainloop()