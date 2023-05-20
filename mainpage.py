from tkinter import *
from views import *  #菜单栏对应的各个子页面
 
class MainPage(object):
	def __init__(self, master=None):
		self.root = master #定义内部变量root
		self.root.geometry('%dx%d' % (600, 550)) #设置窗口大小
		self.root.resizable(width=False, height=False) #宽不可变, 高不可变,默认为True
		self.createPage()
 
	def createPage(self):
		self.replacePage = ReplaceFrame(self.root) # 创建不同Frame
		self.summaryPage = SummaryFrame(self.root)
		self.AddPage = AddFrame(self.root)
		self.aboutPage = AboutFrame(self.root)
		#self.summaryPage.pack() #默认文件汇总界面
		self.AddPage.pack() #默认替换字符界面
		menubar = Menu(self.root)
		menubar.add_command(label='模板汇总', command = self.summaryData)
		menubar.add_command(label='新增汇总', command = self.addData)		
		menubar.add_command(label='批量替换', command = self.replaceData)
		#menubar.add_command(label='文件汇总', command = self.summaryData)
		menubar.add_command(label='关于', command = self.aboutProg)
		self.root['menu'] = menubar  # 设置菜单栏
 
	def replaceData(self):
		self.replacePage.pack()
		self.summaryPage.pack_forget()
		self.AddPage.pack_forget()
		self.aboutPage.pack_forget()
 
	def summaryData(self):
		self.replacePage.pack_forget()
		self.summaryPage.pack()
		self.AddPage.pack_forget()
		self.aboutPage.pack_forget()

	def addData(self):
		self.replacePage.pack_forget()
		self.summaryPage.pack_forget()
		self.AddPage.pack()
		self.aboutPage.pack_forget()
		
	def aboutProg(self):
		self.replacePage.pack_forget()
		self.summaryPage.pack_forget()
		self.AddPage.pack_forget()
		self.aboutPage.pack()
