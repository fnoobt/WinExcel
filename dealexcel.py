import os,shutil
import re
import win32com.client
import tkinter
from easyexcel import *
import time
from win32timezone import *
from tkinter.messagebox import *
#from win32timezone import *

class RepExcel(object):
	def __init__(self, master, logList):
		self.root = master #定义内部变量root
		self.loglist = logList
		self.fileList = []	#文件夹内文件名列表

	def set_dirpath(self, indirpath):
		indirpath = r'%s'%indirpath
		self.dirpath = indirpath.replace('/', '\\')
		self.wrlog("选择的目录：" + self.dirpath)

	def set_srcstr(self, inscrstr):
		self.srcstr = inscrstr
		self.wrlog("查找的目标：'" + self.srcstr + "'")
		
	def set_desstr(self, indesstr):
		self.desstr = indesstr
		self.wrlog("替换为：'" + self.desstr + "'")

	def wrlog(self, logMsg):
		logMsg = time.strftime("%Y-%m-%d %H:%M:%S ",time.localtime())+ logMsg + "\n"
		self.loglist.insert(tkinter.END, logMsg)
		self.root.update_idletasks()
		print(logMsg)
		
	#获取文件夹下所有文件路径
	def listFiles(self):
		#返回一个三元组，遍历的路径、当前遍历路径下的目录、当前遍历目录下的文件名
		for root,dirs,files in os.walk(self.dirpath):
			for fileObje in files:
				if not fileObje.startswith("~$") and  fileObje.endswith((".xlsx", ".xls")) :
					self.fileList.append(os.path.join(root,fileObje))
		
	def stReplace(self):
		'''替换字符'''
		self.listFiles()
		self.wrlog("发现excel文件共：" + str(len(self.fileList)) + "个")
		m_excel = win32com.client.DispatchEx('Excel.Application')
		#隐藏窗口
		m_excel.Visible = 0
		#不显示警告，万一未查找到，不会警告
		m_excel.DisplayAlerts = 0
		'''#m_book = EasyExcel()'''
		for fileObj in self.fileList:
			filename = os.path.basename(fileObj)
			self.wrlog("准备处理文件：" + filename)
			m_book = m_excel.Workbooks.Open(fileObj)
			sheet_max = m_book.Worksheets.Count
			sheet_list = range(1, sheet_max + 1)
			for i in sheet_list:
				m_sheet = m_book.Worksheets(i)
				'''
				#不知为何，用Selection替换字符无法使用
				m_sheet.Select
				m_sheet.Activate
				state = m_excel.Selection.Replace(self.srcstr, self.desstr)
				'''
				m_sheet.Usedrange.Replace(self.srcstr, self.desstr)	
			self.wrlog("文件：" + filename + "替换完成")
			m_book.Save()
			m_book.Close(SaveChanges=1)
		m_excel.quit()
		self.wrlog("文件处理完成")

class AddExcel(object):
#新增表格汇总
	def __init__(self, master, logList, titlefilepath, tempdirpath, desfile):
		self.root = master #定义内部变量root
		self.loglist = logList
		self.fileList = []	#文件夹内文件名列表
		self.rowpointer = 1	#行指针
		self.colpointer = 1	#列指针
		self.titlefilepath = self.deal_path(titlefilepath)
		self.tempdirpath = self.deal_path(tempdirpath)
		self.sum_root = os.path.dirname(self.titlefilepath)
		titletempname, self.type = os.path.splitext(self.titlefilepath)	#按.拆分文件名
				
		self.desfile = desfile
		self.desfilepath = self.deal_filename(desfile)
		
		self.wrlog("选择的标题：" + self.titlefilepath)
		self.wrlog("需汇总的文件夹：" + self.tempdirpath)
		self.wrlog("汇总的文件路径：" + self.desfilepath)
		self.stAdd()
	
	def deal_path(self, inpath):
		inpath = r'%s'%inpath
		outpath = inpath.replace('/', '\\')
		return( outpath )
		
	def deal_filename(self, infile):
		if not infile.endswith(self.type) :
			outfile = infile + self.type
		outfile = os.path.join(self.sum_root,outfile)
		if os.path.exists(outfile):
			infile = self.desfile + "(" + str(self.filenum) + ")"
			self.filenum = self.filenum + 1
			return self.deal_filename(infile)
		else:
			return outfile
			
	def wrlog(self, logMsg):
		logMsg = time.strftime("%Y-%m-%d %H:%M:%S ",time.localtime())+ logMsg + "\n"
		self.loglist.insert(tkinter.END, logMsg)
		self.root.update_idletasks()
		print(logMsg)
	
	def infolog(self, logMsg):
		logMsg = logMsg + "\n"
		self.loglist.insert(tkinter.END, logMsg)
		self.root.update_idletasks()
		print(logMsg)

	#获取文件夹下所有文件路径
	def listFiles(self):
		#返回一个三元组，遍历的路径、当前遍历路径下的目录、当前遍历目录下的文件名
		for root,dirs,files in os.walk(self.tempdirpath):
			for fileObje in files:
				if not fileObje.startswith("~$") and  fileObje.endswith((".xlsx", ".xls")) :
					self.fileList.append(os.path.join(root,fileObje))

	def copyFile(self, srcfile):
		if not os.path.isfile(srcfile):
			print ("%s not exist!"%(srcfile))
		else:
			shutil.copyfile(srcfile,self.desfilepath)      #复制文件
			self.wrlog("复制源文件%s -> %s"%( srcfile,self.desfilepath))				
					
	def dealTempTable(self, file):
	#处理需汇总的表
		for sheet_count in range(1, self.temp_book.getSheetCount()+1):
		#遍历sheet
			#row行，Column列
			#m_sheet = m_book.getSheet(i)
			des_sheet = self.des_book.getSheet(1)
			temp_sheet = self.temp_book.getSheet(sheet_count)
			
			tempRowMax = self.temp_book.getUseRow(sheet_count)
			tempColMax = self.temp_book.getUseCol(sheet_count)
			count_flag = False
			
			if(tempRowMax > self.titleRowMax and tempColMax >= self.titleColMax):
				for row_count in range(1, self.titleColMax + 1):
					des_title = self.des_book.getRangeValue(1, 1, row_count, self.titleRowMax, row_count)
					#print(des_title)
					break_flag = False
					for x in range(1, 7):
						if break_flag:
							break
						for y in range(1, tempColMax + 1):
							temp_title = self.temp_book.getRangeValue(sheet_count, x, y, x + self.titleRowMax - 1, y)
							if des_title == temp_title:
								if self.rowpointer == 1:
									self.rowpointer = self.titleRowMax + 1
								if x + self.titleRowMax > tempRowMax:
									break
								elif x + self.titleRowMax == tempRowMax:
									temp_cell = self.temp_book.getCellValue(sheet_count, tempRowMax, y)
									self.des_book.setCellValue(1, self.rowpointer, row_count, temp_cell)
								else:
									temp_rang = self.temp_book.getRangeValue(sheet_count, x + self.titleRowMax, y, tempRowMax, y)
									self.des_book.setRangeValue(1, self.rowpointer, row_count, temp_rang)
								rowspan = tempRowMax - x - self.titleRowMax + 1
								for i in range(self.rowpointer, self.rowpointer + rowspan):
									self.des_book.setCellValue(1, i, self.titleColMax + 1, file)
								break_flag = True
								count_flag = True
								break
				if count_flag:
					self.rowpointer = self.rowpointer + rowspan
			else:
				print("第sheet" + str(sheet_count) + "不包含想要的数据")
			
	def stAdd(self):
		self.listFiles()
		self.copyFile(self.titlefilepath)
		row_list = []
		col_list = []
		self.wrlog("发现excel模板文件共：" + str(len(self.fileList)) + "个")
		self.des_book = EasyExcel(1)
		self.temp_book = EasyExcel(1)
		self.des_book.open(self.desfilepath)
		self.titleRowMax = self.des_book.getUseRow(1)
		self.titleColMax = self.des_book.getUseCol(1)
		self.des_book.setCellValue(1, self.rowpointer, self.titleColMax + 1, "来源")
		for fileObj in self.fileList:
			sfilename = os.path.basename(self.titlefilepath)
			self.filename = os.path.basename(fileObj)
			self.wrlog("准备处理文件：" + self.filename)
			if (self.filename == sfilename):
				mess = "文件：" + self.filename + "，与模板文件同名，可能为模板文件，将停止此文件的操作"
				#showinfo(title='警告', message = mess)
				self.wrlog("警告：\n！！！！！！！！！！！！！！！！\n>>>>" + mess + "<<<<\n！！！！！！！！！！！！！！！！")
				continue
			print("行指针：" + str(self.rowpointer))
			self.temp_book.open(fileObj)
			
			self.dealTempTable(fileObj)
			self.des_book.save()
			self.temp_book.closeFile()
			self.wrlog("文件：" + self.filename + "处理完成")
		self.des_book.closeFile()
		self.des_book.quitApp()
		self.temp_book.quitApp()
		self.wrlog("恭喜，检索到的文件全部处理完成")
		
class SumExcel(object):
	def __init__(self, master, logList, srcfile, srcdir, desfile):
		self.root = master #定义内部变量root
		self.loglist = logList
		self.fileList = []	#文件夹内文件名列表
		self.filenum = 1
		self.srcfilepath = self.deal_path(srcfile)
		self.srcdirpath = self.deal_path(srcdir)
		self.desfile = desfile
		self.desfilepath = self.deal_filename(desfile)
		self.logpath = os.path.join(self.src_root,"runlog.xls")
		self.wrlog("选择的文件：" + self.srcfilepath)
		self.wrlog("选择的目录：" + self.srcdirpath)
		self.wrlog("汇总的文件路径：" + self.desfilepath)
		self.stSummary()
	
	def deal_path(self, inpath):
		inpath = r'%s'%inpath
		outpath = inpath.replace('/', '\\')
		return( outpath )
	
	def deal_filename(self, infile):
		self.src_root = os.path.dirname(self.srcfilepath)
		#suffix_name = os.path.splitext(self.srcfilepath)[1] 
		self.srcfilename,type = os.path.splitext(self.srcfilepath)
		if not infile.endswith(type) :
			outfile = infile + type
		outfile = os.path.join(self.src_root,outfile)
		if os.path.exists(outfile):
			infile = self.desfile + "(" + str(self.filenum) + ")"
			self.filenum = self.filenum + 1
			return self.deal_filename(infile)
		else:
			return outfile

	def wrlog(self, logMsg):
		logMsg = time.strftime("%Y-%m-%d %H:%M:%S ",time.localtime())+ logMsg + "\n"
		self.loglist.insert(tkinter.END, logMsg)
		self.root.update_idletasks()
		print(logMsg)
	
	def infolog(self, logMsg):
		logMsg = logMsg + "\n"
		self.loglist.insert(tkinter.END, logMsg)
		self.root.update_idletasks()
		print(logMsg)

	#获取文件夹下所有文件路径
	def listFiles(self):
		#返回一个三元组，遍历的路径、当前遍历路径下的目录、当前遍历目录下的文件名
		for root,dirs,files in os.walk(self.srcdirpath):
			for fileObje in files:
				if not fileObje.startswith("~$") and  fileObje.endswith((".xlsx", ".xls")) :
					self.fileList.append(os.path.join(root,fileObje))

	def copyFile(self):
		if not os.path.isfile(self.srcfilepath):
			print ("%s not exist!"%(self.srcfilepath))
		else:
			fpath,fname = os.path.split(self.desfilepath)    #分离文件名和路径
			if not os.path.exists(fpath):
				os.makedirs(fpath)                #创建路径
			shutil.copyfile(self.srcfilepath,self.desfilepath)      #复制文件
			self.wrlog("复制源文件%s -> %s"%( self.srcfilepath,self.desfilepath))
			
	def writeLog(self, count, sumtime, src_book, bra_book, sheet, row, col, src_content, bra_content):
	#写入的行数，，，时间，源文件，来源文件，工作表，行，列，原内容，修改的内容
		self.log_book.setCellValue(1, count, 1, sumtime)
		self.log_book.setCellValue(1, count, 2, src_book)
		self.log_book.setCellValue(1, count, 3, bra_book)
		self.log_book.setCellValue(1, count, 4, sheet)
		self.log_book.setCellValue(1, count, 5, row)
		self.log_book.setCellValue(1, count, 6, col)
		self.log_book.setCellValue(1, count, 7, src_content)
		self.log_book.setCellValue(1, count, 8, bra_content)
	
	def stSummary(self):
		self.listFiles()
		self.copyFile()
		row_list = []
		col_list = []
		firstflag = True
		self.wrlog("发现excel文件共：" + str(len(self.fileList)) + "个")
		src_book = EasyExcel(1)
		des_book = EasyExcel(1)
		bra_book = EasyExcel(1)
		self.log_book = EasyExcel(1)
		src_book.open(self.srcfilepath)
		des_book.open(self.desfilepath)
		self.log_book.open(self.logpath)
		log_count = self.log_book.getUseRow(1)
		self.writeLog(1, "执行时间", "模板文件名", "汇总的文件名", "工作表名", "修改的行号", "修改的列号", "模板的内容", "汇总的内容")
		sheet_list = src_book.getSheetNameList()
		sfilename = os.path.basename(self.srcfilepath)
		lfilename = os.path.basename(self.logpath)
		dfilename = os.path.basename(self.desfilepath)
		for fileObj in self.fileList:
			filename = os.path.basename(fileObj)
			self.wrlog("准备处理文件：" + filename)
			if (filename == sfilename):
				mess = "文件：" + filename + "，与模板文件同名，可能为模板文件，将停止此文件的操作"
				#showinfo(title='警告', message = mess)
				self.wrlog("警告：\n！！！！！！！！！！！！！！！！\n>>>>" + mess + "<<<<\n！！！！！！！！！！！！！！！！")
				continue
			elif (filename == lfilename):
				mess = "文件：" + filename + "，与日志文件同名，可能为日志文件，将停止此文件的操作"
				#showinfo(title='警告', message = mess)
				self.wrlog("警告：\n！！！！！！！！！！！！！！！！\n>>>>" + mess + "<<<<\n！！！！！！！！！！！！！！！！")
				continue
			elif (filename == dfilename):
				mess = "文件：" + filename + "，与汇总文件同名，可能为汇总文件，将停止此文件的操作"
				#showinfo(title='警告', message = mess)
				self.wrlog("警告：\n！！！！！！！！！！！！！！！！\n>>>>" + mess + "<<<<\n！！！！！！！！！！！！！！！！")
				continue	
			bra_book.open(fileObj)
			for sheet_name in sheet_list:
				#row行，Column列
				#m_sheet = m_book.getSheet(i)
				src_sheet = src_book.getSheetIndexByName(sheet_name)
				bar_sheet = bra_book.getSheetIndexByName(sheet_name)
				if(bar_sheet == None):
					mess = "文件：" + filename + "，无法检索到工作表：" + sheet_name + "，将停止此文件的操作"
					#showinfo(title='警告', message = mess)
					self.wrlog("警告：\n！！！！！！！！！！！！！！！！\n>>>>" + mess + "<<<<\n！！！！！！！！！！！！！！！！")
					break
				print(bar_sheet)
				if firstflag:
					row_max = src_book.getUseRow(src_sheet)
					col_max = src_book.getUseCol(src_sheet)
					row_list.insert(src_sheet, row_max)
					col_list.insert(src_sheet, col_max)
					des_book.setSheetName(src_sheet, sheet_name)
				else:
					row_max = row_list[src_sheet - 1]
					col_max = col_list[src_sheet - 1]
				print("在用的最大行" + str(row_max) + ",最大列" +  str(col_max))
				if(row_max > 0 and col_max > 0):
					for x in range(1, row_max + 1):
						for y in range(1, col_max + 1):
							src_content = src_book.getCellValue(src_sheet, x, y)
							bra_content = bra_book.getCellValue(bar_sheet, x, y)
							if src_content != bra_content:
								des_book.setCellValue(src_sheet, x, y, bra_content)
								log_count = log_count + 1
								sumtime = time.strftime("%Y-%m-%d %H:%M:%S ",time.localtime())
								self.writeLog(log_count, sumtime, self.srcfilename, fileObj, sheet_name, x, y, src_content, bra_content)
								#self.infolog("更改行" + str(x) + ",列" +  str(y) + ",内容更改为>>" +  str(bra_content) + "<<end")
				else:
					print("第sheet" + str(src_sheet) + "无数据")
			des_book.save()
			self.log_book.save(self.logpath)
			bra_book.closeFile()
			if bar_sheet != None:
				firstflag = False
			self.wrlog("文件：" + filename + "处理完成")
		des_book.closeFile()
		self.log_book.closeFile()
		src_book.closeFile()
		des_book.quitApp()
		self.log_book.quitApp()
		bra_book.quitApp()
		src_book.quitApp()
		self.wrlog("恭喜，检索到的文件全部处理完成")
		