'''
Created on Oct 5, 2020

@author: pia.manzon
'''
from openpyxl import workbook #pip install openpyxl
from openpyxl import load_workbook

from openpyxl.styles import Border, Side, Alignment, Font, NamedStyle


class Excel():
    def __init__(self,excelPath):
        self.filePath = excelPath
        self.workbook = load_workbook(excelPath)
        self.sheet = self.workbook.active
        self.myThinBorder = Border(top=Side(border_style='thin', color='000000'))
        self.addStyle()
       
    def addStyle(self):
        if(not self.styleExists("RawDataStyle")):
            self.rawdataStyle = NamedStyle(name = "RawDataStyle", font = Font(name = 'Arial', size =10), alignment = Alignment(horizontal = 'left', vertical = 'bottom'))
            self.workbook.add_named_style(self.rawdataStyle)
        if(not self.styleExists("ChartDataStyle")):
            self.chartDataStyle = NamedStyle(name = "ChartDataStyle", font = Font(name = 'Arial', size =10), alignment = Alignment(horizontal = 'center', vertical = 'bottom'))
            self.workbook.add_named_style(self.chartDataStyle)
      
    def styleExists(self,styleName):
        result = self.workbook.named_styles
         
        if (styleName in result):
            return True
        return False
        
    def getMaxRow(self,columnName):
        maxRow = max((c.row for c in self.sheet[columnName] if c.value is not None))
        return maxRow
    
    def setBorder(self,maxRow,columnStart):
        for index in range(columnStart,columnStart+3):
            self.sheet.cell(row = maxRow, column = index).border = self.myThinBorder 
        
    def inputNewData (self,maxRow,flightDataList,city):       
        newMaxRow = maxRow + 1
        columnStart = 0
        borderRow = newMaxRow

        if (city == "Vancouver"):
            columnStart = 3
                 
        else:
            columnStart = 7
              
#Update Flight Info
#===============================================================================
        for index in range(len(flightDataList)):
            tempCell =  self.sheet.cell(row=newMaxRow, column = columnStart)
            tempCell.value = flightDataList[index]['Date']
            tempCell.style = "RawDataStyle"
            tempCell =  self.sheet.cell(row=newMaxRow, column = (columnStart+1))
            tempCell.value = flightDataList[index]['Flight']
            tempCell.style = "RawDataStyle"
            tempCell = self.sheet.cell(row=newMaxRow, column = (columnStart+2))
            tempCell.value = flightDataList[index]['From']
            tempCell.style = "RawDataStyle"
            newMaxRow+=1   
        self.setBorder(borderRow, columnStart)

#Update Daily Summary
#===============================================================================  
    def updateDailySummary(self,flightVancouverList,flightTorontoList):
        dateToday = flightVancouverList[0]['Date']
        columnStart = 12 
        newMaxRow = self.getMaxRow('l') + 1
        tempCell = self.sheet.cell(row=newMaxRow, column = columnStart)
        tempCell.value = dateToday
        tempCell.style = "RawDataStyle"
        tempCell = self.sheet.cell(row=newMaxRow, column = (columnStart+1))
        tempCell.value = len(flightVancouverList)
        tempCell.style = "ChartDataStyle"
        tempCell =   self.sheet.cell(row=newMaxRow, column = (columnStart+2)) 
        tempCell.value = len(flightTorontoList)
        tempCell.style = "ChartDataStyle"
        
   
    def isAlreadyUpdated(self,dateToday):
        maxRow = self.getMaxRow('C')
        lastDateUpdate = self.sheet.cell(row = maxRow, column = 3).value
        if (dateToday == lastDateUpdate):
            return True
        return False
              
    def saveFile(self):
        self.workbook.save(self.filePath)
        