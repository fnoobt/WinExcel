import os
from time import sleep
import win32com.client
'''
选择任何行的最后一个单元格 ActiveCell.End(xlright).Select
选择任何列的最后单元格 ActiveCell.End(xldown).Select
选择任何行的第一个单元格 ActiveCell.End(xleft).Select
选择任何列的第一个单元格 ActiveCell.End(xlup).Select
'''
class EasyExcel(object):
	'''class of easy to deal with excel'''
	
	def __init__(self, process = 0):
		'''initial excel application'''
		self.m_filename = ''
		self.m_exists = False
		self.process = process
		if self.process:
			self.m_excel = win32com.client.DispatchEx('Excel.Application')
		else:
			self.m_excel = win32com.client.Dispatch('Excel.Application')   #DispatchEx,也可以用Dispatch，前者开启新进程，后者会复用进程中的excel进程
		self.m_excel.DisplayAlerts = False                             #覆盖同名文件时不弹出确认框
		self.m_excel.Visible = False								   #隐藏窗口

	def open(self, filename=''):
		'''open excel file'''
		#检查self.m_book是否存在
		if getattr(self, 'm_book', False):
			self.m_book.Close()
		self.m_filename = filename
		if not self.m_filename :
			self.m_book = self.m_excel.Workbooks.Add()
		elif os.path.exists(self.m_filename):
			self.m_book = self.m_excel.Workbooks.Open(self.m_filename)
		else:
			self.m_book = self.m_excel.Workbooks.Add()

	def reset(self):
		'''reset'''
		self.m_excel = None
		self.m_book = None
		self.m_filename = ''

	def save(self, newfile=''):
		'''save the excel content'''
		assert type(newfile) is str, 'filename must be type string'
		if not newfile:
			self.m_book.Save()
			return
		else:
			self.m_book.SaveAs(newfile)
	
	def closeFile(self):
		'''close the file'''
		self.m_book.Close(SaveChanges=1)
		sleep(0.2)
		self.m_book = None
		self.m_filename = ''
		
	def quitApp(self):
		'''close the application'''
		self.m_excel.Quit()
		sleep(1)
		self.reset()
	
	def close(self):
		'''close the application'''
		self.m_book.Close(SaveChanges=1)
		if self.process:
			self.m_excel.Quit()
		sleep(1)
		self.reset()
	
	def addSheet(self, sheetname=None):
		'''add new sheet, the name of sheet can be modify,but the workbook can't '''
		sht = self.m_book.Worksheets.Add()
		sht.Name = sheetname if sheetname else sht.Name
		return sht
	
	def getSheet(self, sheet=1):
		'''get the sheet object by the sheet index'''
		assert sheet > 0, 'the sheet index must bigger then 0'
		return self.m_book.Worksheets(sheet)
		
	def getSheetByName(self, name):
		'''get the sheet object by the sheet name'''
		for i in range(1, self.getSheetCount()+1):
			sheet = self.getSheet(i)
			if name == sheet.Name:
				return sheet
		return None
	
	def setSheetName(self, sheet, name):
		'''set sheet name'''
		self.getSheet(sheet).Name = name
		
	def getSheetIndexByName(self, name):
		'''get the sheet index by the sheet name'''
		for i in range(1, self.getSheetCount()+1):
			sheet = self.getSheet(i)
			if name == sheet.Name:
				return i
		return None
	
	def getSheetNameList(self):
		'''get the sheet name list object'''
		name_list = []
		for i in range(1, self.getSheetCount()+1):
			sheet = self.getSheet(i)
			name_list.append(sheet.Name)
		return name_list
	
	def getCell(self, sheet=1, row=1, col=1):
		'''get the cell object'''
		assert row>0 and col>0, 'the row and column index must bigger then 0'
		return self.getSheet(sheet).Cells(row, col)
		
	def getRow(self, sheet=1, row=1):
		'''get the row object'''
		assert row>0, 'the row index must bigger then 0'
		return self.getSheet(sheet).Rows(row)
		
	def getCol(self, sheet, col):
		'''get the column object'''
		assert col>0, 'the column index must bigger then 0'
		return self.getSheet(sheet).Columns(col)

	def getRowByName(self, name, row=1):
		'''get the row object'''
		assert row>0, 'the row index must bigger then 0'
		return self.getSheetByName(name).Rows(row)
		
	def getColByName(self, name, col):
		'''get the column object'''
		assert col>0, 'the column index must bigger then 0'
		return self.getSheetByName(name).Columns(col)
	
	def getRange(self, sheet, row1, col1, row2, col2):
		'''get the range object'''
		sht = self.getSheet(sheet)
		return sht.Range(self.getCell(sheet, row1, col1), self.getCell(sheet, row2, col2))
	
	def getCellValue(self, sheet, row, col):
		'''Get value of one cell'''
		return self.getCell(sheet,row, col).Value
		
	def setCellValue(self, sheet, row, col, value):
		'''set value of one cell'''
		self.getCell(sheet, row, col).Value = value
		
	def getRowValue(self, sheet, row):
		'''get the row values'''
		return self.getRow(sheet, row).Value
		
	def setRowValue(self, sheet, row, values):
		'''set the row values'''
		self.getRow(sheet, row).Value = values
	
	def getRowValueByName(self, name, row):
		'''get the row values'''
		return self.getRowByName(name, row).Value
	
	def getColValue(self, sheet, col):
		'''get the row values'''
		return self.getCol(sheet, col).Value
		
	def setColValue(self, sheet, col, values):
		'''set the row values'''
		self.getCol(sheet, col).Value = values
	
	def getColValueByName(self, name, row):
		'''get the row values'''
		return self.getColByName(name, row).Value
	
	def getRangeValue(self, sheet, row1, col1, row2, col2):
		'''return a tuples of tuple)'''
		return self.getRange(sheet, row1, col1, row2, col2).Value
	
	def setRangeValue(self, sheet, row1, col1, data):
		'''set the range values'''
		row2 = row1 + len(data) - 1
		col2 = col1 + len(data[0]) - 1
		range = self.getRange(sheet, row1, col1, row2, col2)
		range.Clear()
		range.Value = data
		
	def getSheetCount(self):
		'''get the number of sheet'''
		return self.m_book.Worksheets.Count
	
	def getMaxRow(self, sheet):
		'''get the max row number, not the count of used row number'''
		return self.getSheet(sheet).Rows.Count
		
	def getMaxCol(self, sheet):
		'''get the max col number, not the count of used col number'''
		return self.getSheet(sheet).Columns.Count
	
	def getUseRow(self, sheet):
		'''get the use row number'''
		row_max = self.getSheet(sheet).Usedrange.Rows.Count
		col_max = self.getSheet(sheet).Usedrange.Columns.Count
		#print("excel使用过的最大行数" + str(row_max))
		while row_max > 1 :
			#row_list = self.getRowValue(sheet, row_max)[0]
			#二维列表，获取第一行
			row_list = self.getRangeValue(sheet, row_max, 1, row_max, col_max)
			if (col_max > 1):
				del_list = list(filter(None, row_list[0]))
				if (del_list == []):
					row_max = row_max - 1
				else:
					return row_max
			else:
				if(row_list == None):
					row_max = row_max - 1
				else:
					break
		return row_max
		
	def getUseCol(self, sheet):
		'''get the use col number'''
		row_max = self.getSheet(sheet).Usedrange.Rows.Count
		col_max = self.getSheet(sheet).Usedrange.Columns.Count
		#print("excel使用过的最大列数" + str(col_max))
		while col_max >1:
			#col_list = self.getColValue(sheet, col_max)
			#二维列表，获取第一列
			col_list = self.getRangeValue(sheet, 1, col_max, row_max, col_max)
			if (row_max > 1):
				#col_list = map(str, col_list)
				del_list = [i[0] for i in col_list]
				del_list = list(filter(None, del_list))
				if(del_list == []):
					col_max = col_max - 1
				else:
					return col_max
			else:
				if(col_list == None):
					col_max = col_max - 1
				else:
					break
		return col_max
	
	def clearCell(self, sheet, row, col):
		'''clear the content of the cell'''
		self.getCell(sheet,row,col).Clear()
		
	def deleteCell(self, sheet, row, col):
		'''delete the cell'''
		self.getCell(sheet, row, col).Delete()
		
	def clearRow(self, sheet, row):
		'''clear the content of the row'''
		self.getRow(sheet, row).Clear()
		
	def deleteRow(self, sheet, row):
		'''delete the row'''
		self.getRow(sheet, row).Delete()
		
	def clearCol(self, sheet, col):
		'''clear the col'''
		self.getCol(sheet, col).Clear()
		
	def deleteCol(self, sheet, col):
		'''delete the col'''
		self.getCol(sheet, col).Delete()
		
	def clearSheet(self, sheet):
		'''clear the hole sheet'''
		self.getSheet(sheet).Clear()
		
	def deleteSheet(self, sheet):
		'''delete the hole sheet'''
		self.getSheet(sheet).Delete()
	
	def deleteRows(self, sheet, fromRow, count=1):
		'''delete count rows of the sheet'''
		maxRow = self.getMaxRow(sheet)
		maxCol = self.getMaxCol(sheet)
		endRow = fromRow+count-1
		if fromRow > maxRow or endRow < 1:
			return
		self.getRange(sheet, fromRow, 1, endRow, maxCol).Delete()
		
	def deleteCols(self, sheet, fromCol, count=1):
		'''delete count cols of the sheet'''
		maxRow = self.getMaxRow(sheet)
		maxCol = self.getMaxCol(sheet)
		endCol = fromCol + count - 1
		if fromCol > maxCol or endCol < 1:
			return
		self.getRange(sheet, 1, fromCol, maxRow, endCol).Delete()