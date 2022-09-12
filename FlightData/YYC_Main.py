'''
Created on Oct 2, 2020

@author: pia.manzon
'''
import YYC_Webpage
import YYC_Excel

cityName = ""

excelFileOne = YYC_Excel.Excel(r"")
 
 
cityName = "Vancouver"
YYC_Webpage.clickArrivalButton()
YYC_Webpage.searchCity(cityName)
YYC_Webpage.clickLoadFlightsButton()
flightVancouverList = YYC_Webpage.getFlightSched() #get Vancouver data
 
#Check if date already exist in the file: no update needed 
maxRowVancouver = excelFileOne.getMaxRow('C')

if (excelFileOne.isAlreadyUpdated(flightVancouverList[0]['Date'])):
    YYC_Webpage.closeBrowser()
    print("Already Updated")
    exit()
excelFileOne.inputNewData(maxRowVancouver, flightVancouverList,cityName)
  
cityName = "Toronto"
YYC_Webpage.searchCity(cityName)
YYC_Webpage.clickLoadFlightsButton()
flightTorontoList = YYC_Webpage.getFlightSched() #get Toronto data
   
YYC_Webpage.closeBrowser()
  
maxRowToronto = excelFileOne.getMaxRow('G')
excelFileOne.inputNewData(maxRowToronto, flightTorontoList,cityName)
excelFileOne.updateDailySummary(flightVancouverList,flightTorontoList)
#excelFileOne.updateWeeklySummary()
excelFileOne.saveFile()
print("Done saving")

