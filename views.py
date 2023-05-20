from tkinter import *
from tkinter.messagebox import *
from tkinter.filedialog import askdirectory
from tkinter.filedialog import askopenfilename
import re
from dealexcel import * #替换字符的处理函数
from win32timezone import *
from tkinter.scrolledtext import ScrolledText

class ReplaceFrame(Frame): # 继承Frame类
	def __init__(self, master=None):
		Frame.__init__(self, master)
		self.root = master #定义内部变量root
		self.dirpath = StringVar()
		self.srcstr = StringVar()
		self.desstr = StringVar()
		#self.dirpath.set(r"D:/PythonProject/pytkinter/winexcel/excelfile")
		#self.srcstr.set("安化")
		#self.desstr.set("一个")
		self.createPage()
		
	def createPage(self):
		#grid窗体布局，row表示行，column表示列,ipadx、ipady组件的内部间隙,padx、pady外部间隔距离,columnspan跨越的列数,rowspan跨越的行数,sticky对齐方式(N、E、S、W)
		Label(self, text='EXCEL字符批量替换').grid(row=0, column=1, pady=10)
		Label(self, text = '文件夹路径:').grid(row=1, column=0)
		Entry(self, textvariable = self.dirpath).grid(row=1, column=1)
		Button(self, text = "选择文件夹", command = self.selectDirPath).grid(row = 1, column = 2)
		
		Label(self, text = '查找目标：').grid(row=2, column=0)
		Entry(self, textvariable=self.srcstr).grid(row=2, column=1)
		Label(self, text = '替换为: ').grid(row=3, column=0)
		Entry(self, textvariable=self.desstr).grid(row=3, column=1)
		Button(self, text='开始替换', command = self.startReplace).grid(row=5, column=1, pady=10)
		Label(self, text = '运行状态').grid(row=6, column=0)
		self.logList = ScrolledText(self)
		#self.logList = Text(self)
		self.logList.tag_config("greencolor", foreground='#008C00')
		self.logList.grid(row=7, column=0, columnspan=3)
		#0.0是0行0列到END，表示全部，END表示插入末端
		#logList.insert(END,txtMsg.get('0.0',END))

	'''
	**打开一个文件：**askopenfilename()
	**打开一组文件：**askopenfilenames()
	**保存文件：**asksaveasfilename()
	'''
	def selectDirPath(self):
		#返回文件夹路径
		dirpath_ = askdirectory()
		self.dirpath.set(dirpath_)
	
	def delEnter(self, str):
		newstr = str.strip("\r\n\t")
		return newstr		

	def startReplace(self):
		self.logList.delete('0.0','end')
		getdirpath = self.dirpath.get()
		getsrcstr = self.delEnter(self.srcstr.get())
		getdesstr = self.delEnter(self.desstr.get())
		if(getdirpath == ""):
			showinfo(title='错误', message='未选择文件夹')
		elif(getsrcstr == "" ):
			showinfo(title='错误', message='查找的字符为空')
		else:
			try:
				rex = RepExcel(self.root, self.logList)
				rex.set_dirpath(getdirpath)
				rex.set_srcstr(getsrcstr)
				rex.set_desstr(getdesstr)
				rex.stReplace()
			except:
				showinfo(title='错误', message='程序出现未知错误，请退出程序，并关闭所有excel文件。')

class SummaryFrame(Frame): # 继承Frame类
	def __init__(self, master=None):
		Frame.__init__(self, master)
		self.root = master #定义内部变量root
		self.srcfilepath = StringVar()
		self.srcdirpath = StringVar()
		self.desfilename = StringVar()
		#self.dirpath.set(r"D:/PythonProject/pytkinter/winexcel/excelfile"))
		self.createPage()
		
	def createPage(self):
		#grid窗体布局，row表示行，column表示列,ipadx、ipady组件的内部间隙,padx、pady外部间隔距离,columnspan跨越的列数,rowspan跨越的行数,sticky对齐方式(N、E、S、W)
		Label(self, text='EXCEL文件模板汇总').grid(row=0, column=1, pady=10)
		Label(self, text = '选择模板文件路径:').grid(row=1, column=0)
		Entry(self, textvariable = self.srcfilepath).grid(row=1, column=1)
		Button(self, text = " 选择文件 ", command = self.selectFilePath).grid(row = 1, column = 2)
		
		Label(self, text = '选择需汇总文件夹:').grid(row=2, column=0)
		Entry(self, textvariable = self.srcdirpath).grid(row=2, column=1)
		Button(self, text = "选择文件夹", command = self.selectDirPath).grid(row = 2, column = 2)
		
		Label(self, text = '汇总后的文件名：').grid(row=3, column=0)
		Entry(self, textvariable=self.desfilename).grid(row=3, column=1)
		Button(self, text='开始汇总', command = self.startSummary).grid(row=4, column=1, pady=10)
		Label(self, text = '运行状态').grid(row=5, column=0)
		self.logList = ScrolledText(self)
		#self.logList = Text(self)
		self.logList.tag_config("greencolor", foreground='#008C00')
		self.logList.grid(row=7, column=0, columnspan=3)
		#0.0是0行0列到END，表示全部，END表示插入末端
		#logList.insert(END,txtMsg.get('0.0',END))

	'''
	**打开一个文件：**askopenfilename()
	**打开一组文件：**askopenfilenames()
	**保存文件：**asksaveasfilename()
	'''
	def selectDirPath(self):
		#返回文件夹路径
		dirpath_ = askdirectory()
		self.srcdirpath.set(dirpath_)
		
	def selectFilePath(self):
		#返回文件路径
		filepath_ = askopenfilename()
		self.srcfilepath.set(filepath_)
		
	def checkName(self, name=None):
		if name is None:
			print("name is None!")
			return False
		reg = re.compile(r'[\\/:*?"<>|\r\n]+')
		valid_name = reg.findall(name)
		if valid_name:
			return False
		else:
			return name
		'''
		if valid_name:
		for nv in valid_name:
		name = name.replace(nv, "_")
		return name
		'''

	def startSummary(self):
		self.logList.delete('0.0','end')
		getsrcfilepath = self.srcfilepath.get()
		getsrcdirpath = self.srcdirpath.get()
		getdesfilename = self.checkName(self.desfilename.get())
		if(getsrcfilepath == ""):
			showinfo(title='错误', message='未选择源文件！')
		elif(getsrcdirpath == "" ):
			showinfo(title='错误', message='未选择文件夹！')
		elif not getdesfilename:
			showinfo(title='提示', message='未设置汇总后的文件名,或设置的文件名格式错误！')
		else:
			#rex = SumExcel(self.root, self.logList, getsrcfilepath, getsrcdirpath, getdesfilename)
			try:
				rex = SumExcel(self.root, self.logList, getsrcfilepath, getsrcdirpath, getdesfilename)
			except:
				showinfo(title='错误', message='程序出现未知错误，请退出程序，并关闭所有excel文件。')


class AddFrame(Frame): # 继承Frame类
	def __init__(self, master=None):
		Frame.__init__(self, master)
		self.root = master #定义内部变量root
		self.srcfilepath = StringVar()
		self.srcdirpath = StringVar()
		self.desfilename = StringVar()
		#self.dirpath.set(r"D:/PythonProject/pytkinter/winexcel/excelfile"))
		self.createPage()
		
	def createPage(self):
		#grid窗体布局，row表示行，column表示列,ipadx、ipady组件的内部间隙,padx、pady外部间隔距离,columnspan跨越的列数,rowspan跨越的行数,sticky对齐方式(N、E、S、W)
		Label(self, text='EXCEL文件新增汇总').grid(row=0, column=1, pady=10)
		Label(self, text = '选择标题文件：:').grid(row=1, column=0)
		Entry(self, textvariable = self.srcfilepath).grid(row=1, column=1)
		Button(self, text = " 选择文件 ", command = self.selectFilePath).grid(row = 1, column = 2)
		
		Label(self, text = '选择需汇总的文件夹：:').grid(row=2, column=0)
		Entry(self, textvariable = self.srcdirpath).grid(row=2, column=1)
		Button(self, text = "选择文件夹", command = self.selectDirPath).grid(row = 2, column = 2)
		
		Label(self, text = '汇总后的文件名：').grid(row=3, column=0)
		Entry(self, textvariable=self.desfilename).grid(row=3, column=1)
		Button(self, text='开始汇总', command = self.startSummary).grid(row=4, column=1, pady=10)
		Label(self, text = '运行状态').grid(row=5, column=0)
		self.logList = ScrolledText(self)
		#self.logList = Text(self)
		self.logList.tag_config("greencolor", foreground='#008C00')
		self.logList.grid(row=7, column=0, columnspan=3)
		#0.0是0行0列到END，表示全部，END表示插入末端
		#logList.insert(END,txtMsg.get('0.0',END))

	'''
	**打开一个文件：**askopenfilename()
	**打开一组文件：**askopenfilenames()
	**保存文件：**asksaveasfilename()
	'''
	def selectDirPath(self):
		#返回文件夹路径
		dirpath_ = askdirectory()
		self.srcdirpath.set(dirpath_)
		
	def selectFilePath(self):
		#返回文件路径
		filepath_ = askopenfilename()
		self.srcfilepath.set(filepath_)
		
	def checkName(self, name=None):
		if name is None:
			print("name is None!")
			return False
		reg = re.compile(r'[\\/:*?"<>|\r\n]+')
		valid_name = reg.findall(name)
		if valid_name:
			return False
		else:
			return name
		'''
		if valid_name:
		for nv in valid_name:
		name = name.replace(nv, "_")
		return name
		'''

	def startSummary(self):
		self.logList.delete('0.0','end')
		getsrcfilepath = self.srcfilepath.get()
		getsrcdirpath = self.srcdirpath.get()
		getdesfilename = self.checkName(self.desfilename.get())
		if(getsrcfilepath == ""):
			showinfo(title='错误', message='未选择源文件！')
		elif(getsrcdirpath == "" ):
			showinfo(title='错误', message='未选择文件夹！')
		elif not getdesfilename:
			showinfo(title='提示', message='未设置汇总后的文件名,或设置的文件名格式错误！')
		else:
			#rex = AddExcel(self.root, self.logList, getsrcfilepath, getsrcdirpath, getdesfilename)
			try:
				rex = AddExcel(self.root, self.logList, getsrcfilepath, getsrcdirpath, getdesfilename)
			except:
				showinfo(title='错误', message='程序出现未知错误，请退出程序，并关闭所有excel文件。')


class AboutFrame(Frame): # 继承Frame类
	def __init__(self, master=None):
		Frame.__init__(self, master)
		self.root = master #定义内部变量root
		self.createPage()
 
	def createPage(self):
		Label(self, text='关于软件').pack()
		Label(self, text='version:1.3').pack()
		shuoming = "使用须知：\n\
1.软件依赖Microsoft Office，如果您使用的是WPS，软件可能无法正常使用。\n\
2.对于较低版本的Windows系统(如XP)可能存在依赖库缺失的问题，请安装相应库。\n\
3.软件兼容xls和xlsx两种格式的Excel。\n\
4.软件在操作时不会改变原有的Excel格式(字体、颜色等)。\n\
5.使用本软件请确保没有打开需要操作的Excel文件。\n\
6.使用本软件前，请提前备份原文件，本软件不确保能完全按照使用者意愿运行。\n\
7.因使用本软件导致的数据丢失，文件损坏等问题由使用者自行承担。\n\
8.如有问题或建议可与我联系，邮箱：my_hb@139.com。\n"
		Label(self, text = shuoming, justify=LEFT).pack(anchor=W)
		#noteText = Text(self)
		#noteText.tag_config("greencolor", foreground='#008C00')
		#noteText.insert(END,shuoming,'greencolor')
		#noteText.pack(anchor=W)
		Label(self, text = "Author:haibo").pack(anchor=SE)