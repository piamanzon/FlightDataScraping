'''
Created on Oct 2, 2020

@author: pia.manzon
'''
import WebPageFunc
import ExcelFunc

cityName = ""
#excelFileOne = ExcelFunc.Excel(r"\\petrochina-ca.com\data\HOME\pia.manzon\YYC_FlightData.xlsx")
excelFileOne = ExcelFunc.Excel(r"\\petrochina-ca.com\data\CANADA\CANADA\Market Data\Oil\Flight Data\YYC Flight Arrivals.xlsx")


cityName = "Vancouver"
WebPageFunc.clickArrivalButton()
WebPageFunc.searchCity(cityName)
WebPageFunc.clickLoadFlightsButton()
flightVancouverList = WebPageFunc.getFlightSched() #get Vancouver data
 
# Check if date already exist in the file: no update needed 
maxRowVancouver = excelFileOne.getMaxRow('C')
if (excelFileOne.isAlreadyUpdated(flightVancouverList[0]['Date'])):
    exit
excelFileOne.inputNewData(maxRowVancouver, flightVancouverList,cityName)
 
cityName = "Toronto"
WebPageFunc.searchCity(cityName)
WebPageFunc.clickLoadFlightsButton()
flightTorontoList = WebPageFunc.getFlightSched() #get Toronto data
  
WebPageFunc.closeBrowser()
 
maxRowToronto = excelFileOne.getMaxRow('G')
excelFileOne.inputNewData(maxRowToronto, flightTorontoList,cityName)
excelFileOne.updateDailySummary(flightVancouverList,flightTorontoList)
excelFileOne.saveFile()

