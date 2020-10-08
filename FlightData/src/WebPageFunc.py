'''
Created on Oct 5, 2020

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
yycFlightURL = 'https://www.yyc.com/en-us/travellerinfo/flightinformation/flightschedule.aspx'
driver.get(yycFlightURL) #open website

   
def clickArrivalButton():
    xPath = '//*[@id="btn_Arrivals"]'
    arrivalButton = driver.find_element_by_xpath(xPath)
    arrivalButton.click()    
    
def searchCity(cityName):
    inputCity = driver.find_element_by_id('yyc-searchInput').clear() 
    inputCity = driver.find_element_by_id('yyc-searchInput')   
    inputCity.send_keys(cityName) 
    
def checkPath(xPath):
    try:
        driver.find_element_by_class_name("btn-show-earlierFlights")
    except NoSuchElementException:
        return False
    return True
    
def clickLoadFlightsButton():  
    xPath = '//*[@id="btn_EarlierFlights"]' 
    result = checkPath(xPath) 
    if(result): 
        loadButton = driver.find_element_by_xpath(xPath) 
        loadButton.click() 
    time.sleep(1)
    
def getFlightSched(): 
    flightInfo = []
    dateToday = date.today()
    dateTodayFormatted = dateToday.strftime('%d-%b-%y').lstrip("0").replace(" 0", " ")
                         
    for flightDetails  in driver.find_elements_by_class_name('yyc-flightDetails '):
        flightNumber = flightDetails .find_element_by_xpath('.//div[@class="airlineCodeDetail"]').text #".//" the dot indicates children of the parent class
        city = flightDetails.find_element_by_xpath('.//div[@class="airportDetail"]').text
        flightStatus = flightDetails.find_element_by_xpath('.//div[@class="statusDetail"]').text
        if(str(flightStatus) != ("Cancelled")):
            flightInfo.append({'Date': dateTodayFormatted,'Flight': flightNumber, 'From': city})
    time.sleep(1)
    return flightInfo

def closeBrowser():
    driver.quit()
    