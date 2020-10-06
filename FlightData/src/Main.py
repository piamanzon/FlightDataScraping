'''
Created on Oct 2, 2020

@author: pia.manzon
'''
import WebPageFunc
import ExcelFunc

cityName = ""
excelFileOne = ExcelFunc.Excel(r"\\petrochina-ca.com\data\HOME\pia.manzon\YYC_FlightData.xlsx")

cityName = "Vancouver"
WebPageFunc.clickArrivalButton()
WebPageFunc.searchCity(cityName)
WebPageFunc.clickLoadFlightsButton()
flightVancouverList = WebPageFunc.getFlightSched()


maxRowVancouver = excelFileOne.getMaxRow('C')
excelFileOne.inputNewData(maxRowVancouver, flightVancouverList,cityName)

cityName = "Toronto"
WebPageFunc.searchCity(cityName)

WebPageFunc.clickLoadFlightsButton()

flightTorontoList = WebPageFunc.getFlightSched()


WebPageFunc.closeBrowser()


maxRowToronto = excelFileOne.getMaxRow('G')
 
excelFileOne.inputNewData(maxRowToronto, flightTorontoList,cityName)
excelFileOne.updateDailySummaryVan(flightVancouverList,flightTorontoList)
excelFileOne.createGraph()
print("before saving")
excelFileOne.saveFile()


