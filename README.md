# FlightDataScraping

<!-- TABLE OF CONTENTS -->
## Table of Contents

* [About the Project](#about-the-project)
  * [Built With](#built-with)
* [Getting Started](#getting-started)
  * [Prerequisites](#prerequisites)
  * [Selenium Installation](#selenium-installation)
  * [Code Installation](#code-installation)
* [Usage](#usage)
* [Contact](#contact)


<!-- ABOUT THE PROJECT -->
## About The Project
The goal of this program is to automatically scrape YYC flight data from the Calgary airport's website and store the data to an excel file. Using python, the excel file will be modified to create charts from those data.
The project is written in Python with Selenium.

![Untitled Diagram (4)](https://user-images.githubusercontent.com/49318818/97631453-a1ac3980-19f6-11eb-9d71-e21d2a6d0375.png)

### Built With

* pyodbc
* selenium
* openpyxl

<!-- GETTING STARTED -->
## Getting Started

To get a local copy up and running follow these simple steps.

### Prerequisites

* python (3.8 or latest version)
```sh
https://www.python.org/downloads/release/python-380/
```
### Selenium Installation
1. Make sure that you have google chrome browser downloaded in your computer
2 Check Chrome browser's version 
  * Open browser
  * Click on the Menu icon in the upper right corner of the screen
  * Click on Help, and then About Google Chrome
3. Download ChromeDriver that matches chrome browser's version
```sh
https://chromedriver.chromium.org/downloads
```
4. Install _selenium_ (if not yet installed) using the command window (Win + R)
```sh
cd "Python path" (ex: C:\Users\pia.manzon\AppData\Local\Programs\Python\Python38\Scripts)
pip install selenium
```

### Code Installation

1. Clone the repo
```sh
https://github.com/piamanzon/FlightDataScraping.git
```
2. Install _openpyxl_ (if not yet installed) using the command window (Win + R)
```sh
cd "Python path" (ex: C:\Users\pia.manzon\AppData\Local\Programs\Python\Python38\Scripts)
pip install openpyxl
```

<!-- USAGE EXAMPLES -->
### Usage
Run independently
1. Open command window (Win + R + Enter)
2. cd "path where YYC_Main.py is saved"
3. Enter YYC_Main.py

Run using Window Task Scheduler
1. Open YYCFlightUpdate.bat_
2. Edit the .bat file by first changing the first quoted line with the python path. Change the second quoted line with the path where the YYC_Main.py is saved
###### Note: If the file is saved under a mapped drive (e.g., H drive) you need to use UNC format. To find the UNC format of your drive use keys Win + R, type cmd and hit Enter. Type _net use_ then copy the UNC path.
3. Open Task Scheduler
4. Create Basic Task with the following configuration: 
    * General : _Run whether user is logged on or not, run with highest priveleges_
    * Trigger: Up to the user
    * Actions: Start a program, Program/script: _"path where the .bat file is saved (UNC format)"_
    * Conditions: Uncheck everything
    * Settings: Allow task to be run on demand, Run task as soon as possible, Stop the task if it runs longer than (), If the running task does not end.., Do not start a new                     instance

<!-- CONTACT -->
### Contact

Name: Pia Angelica Manzon
Email: pia.manzon@petrochina-ca.com, piamanzon@gmail.com


Project Link: [https://github.com/piamanzon/FlightDataScraping](https://github.com/piamanzon/FlightDataScraping)
