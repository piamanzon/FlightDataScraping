'''
Created on Oct 5, 2020

@author: pia.manzon

'''
#Selenium Imports
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

#Pandas Import
import pandas as pd

#Date and Time Imports
from datetime import date
import time



#Other imports
import requests
from bs4 import BeautifulSoup
from django.utils.datetime_safe import strftime
from selenium.common.exceptions import NoSuchElementException

#Selenium
chromedriver_path = r'C:\Users\pia.manzon\Downloads\chromedriver_win32/chromedriver.exe'
driver = webdriver.Chrome(executable_path=chromedriver_path) # This will open the Chrome window
yycFlightURL = 'https://www.yyc.com/en-us/travellerinfo/flightinformation/flightschedule.aspx'
driver.get(yycFlightURL) #open website



    
#<button type="button" id="btn_Arrivals" class="active">Arrivals</button>    
def clickArrivalButton():
    xPath = '//*[@id="btn_Arrivals"]'
    arrivalButton = driver.find_element_by_xpath(xPath)
    arrivalButton.click()    
    
#//*[@id="yyc-searchInput"]
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
    
#//*[@id="btn_EarlierFlights"]
def clickLoadFlightsButton():  
    xPath = '//*[@id="btn_EarlierFlights"]' 
    result = checkPath(xPath) 
    print (result)
    if(result): 
        loadButton = driver.find_element_by_xpath(xPath) 
        loadButton.click() 
    time.sleep(1)
    
    

      
 
#//*[@id="flights-container"]/div/div[1]/div[2]/div[2]/div[5]/span/div/div[3]
#<div class="airlineCodeDetail"><!-- react-text: 3807 --> <!-- /react-text --><!-- react-text: 3808 -->WS126<!-- /react-text --><!-- react-text: 3809 -->  <!-- /react-text --></div>
# //*[@id="flights-container"]/div/div[1]/div[2]/div[2]/div[5]/span/div/div[4]
#<div class="airportDetail">Vancouver</div>
#<div class="yyc-flightDetails "><div class="scheduledDetail"><div>22:22</div></div><div class="updatedDetail">&nbsp;</div><div class="airlineCodeDetail"><!-- react-text: 5087 --> <!-- /react-text --><!-- react-text: 5088 -->WS138<!-- /react-text --><!-- react-text: 5089 -->  <!-- /react-text --></div><div class="airportDetail">Vancouver</div><div class="statusDetail">&nbsp;</div><div class="gateDetail">C</div><div class="airlineDetail">WestJet</div><div class="moreInfoDetail"><button type="button" class="btn btn_guide">More Info</button></div><button type="button" class="btn_desktop_NotifyMe">Notify Me</button><div class="buttonDetail"><i class="fa fa-chevron-up btnClosed"></i><i class="fa fa-chevron-down btnOpen"></i></div></div>
def getFlightSched(): 
    flightInfo = []
    dateToday = date.today()
    dateTodayFormatted = dateToday.strftime('%d-%b-%y')
                         
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
    