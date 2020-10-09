'''
Created on Oct 8, 2020

@author: pia.manzon
'''
#Selenium Imports
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException

#Date and Time Imports
from datetime import date
import time


#Selenium
chromedriver_path = r'C:\Users\pia.manzon\Downloads\chromedriver_win32/chromedriver.exe'
driver = webdriver.Chrome(executable_path=chromedriver_path) # This will open the Chrome window
canadaCityURL = 'https://www.prokerala.com/travel/airports/canada/'
#yvrFlightURL = 'https://www.yvr.ca/en/passengers/flights/arriving-flights'
yvrFlightURL = "https://www.flightradar24.com/data/airports/yvr/arrivals"
driver.get(yvrFlightURL) #open website

#===============================================================================
# #<span class="t-small text-muted">Abbotsford</span>
# try:
#     #===========================================================================
#     # temp = driver.find_element_by_class_name("t-small text-muted").text
#     # print(temp)
#     #===========================================================================
#     
#    # result = driver.find_element_by_xpath('//*[@id="wrapper"]/main/div/div[2]/table[2]/tbody/tr[3]/td[2]/span').text
#     result = driver.find_element_by_xpath("//*[contains(text(), 'Tofino')]").text
#     print ("Domestic")
# except NoSuchElementException:
#     
#     driver.quit()
#     print ("Intl") 
#===============================================================================
def clickEarlierFlights():
    time.sleep(1)
    try:
        loadButton = driver.find_element_by_xpath("//button[contains(text(), 'View earlier flights')]")
        if(loadButton):
            print(loadButton)
            loadButton.click() 
            time.sleep(1)
    except NoSuchElementException:
        
        print("do nothing")
        return
#<button class="yvr-flights__button yvr-flights__button--view-earlier yvr-button yvr-button--delta">View earlier flights</button>

#===============================================================================
# def getFlightData():
#     flightInfo = []
#     dateToday = date.today()
#     dateTodayFormatted = dateToday.strftime('%d-%b-%y').lstrip("0").replace(" 0", " ")
#     #===========================================================================
#     # flightDetails = driver.find_elements_by_id("maincontent")
#     # print(flightDetails)
#     #===========================================================================
#  #\33 05994106 > td.yvr-flights__table-cell--revised.notranslate
#     
#     for flightDetails  in driver.find_elements_by_css_selector("td.yvr-flights__table-cell--revised.notranslate"):
#       #  print("fd" & flightDetails)
#         flightNumber = flightDetails.find_element_by_css_selector("td.yvr-table__cell.yvr-flights__flightNumber.notranslate").text
#         #=======================================================================
#         # city = flightDetails.find_element_by_xpath('.//div[@class="airportDetail"]').text
#         # flightStatus = flightDetails.find_element_by_xpath('.//div[@class="statusDetail"]').text
#         # if(str(flightStatus) != ("Cancelled")):
#         #     flightInfo.append({'Date': dateTodayFormatted,'Flight': flightNumber, 'From': city})
#         #=======================================================================
#         print (flightNumber)
#         flightInfo.append({'FlightNum': flightNumber})
#     time.sleep(1)
#     print (flightInfo)
#     return flightInfo
#===============================================================================

#cnt-data-content > div > div.tab-pane.p-l.active > div > aside > div.row.cnt-schedule-table > table > tbody > tr:nth-child(221) > td.p-l-s.cell-flight-number > a
def getFlightData():
    flightInfo = []
    dateToday = date.today()
    dateTodayFormatted = dateToday.strftime('%d-%b-%y').lstrip("0").replace(" 0", " ")
    #===========================================================================
    # flightDetails = driver.find_elements_by_id("maincontent")
    # print(flightDetails)
    #===========================================================================
    for flightDetails  in driver.find_elements_by_css_selector('tr.data-date="Friday, Oct 09"'):
      #  print("fd" & flightDetails)
        flightNumber = flightDetails.find_element_by_css_selector("td.p-l-s.cell-flight-number").text
        print (flightNumber)
        flightInfo.append({'FlightNum': flightNumber})
        