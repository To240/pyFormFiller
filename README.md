# pyFormFiller
This script uses the Selenium and Openpyxl libraries to extract data from a spreadsheet and input it into HTML forms.

I primarily use this to input data into a CRM as inputting data through a web form allows for validation of fields - whereas usually uploading a sheet of data does not.

change the sheet attributes to correspond to those in your excel sheet, and change the xpaths in the driver.find elements to correspond with the HTML form you are interacting with.

Get an xpath locator plugin for the browser to speed up the process of finding xpaths.

driver.implicitly_wait(1) should be used during testing, so that you can make sure that the correct data is being input into fields.


driver = webdriver.Chrome(r'C:\Users\Tomas\Desktop\ChromeDriver.exe')   -> this is the location of ChromeDriver which is responsible for opening an 'automation' version of Google Chrome.
You may also need to add ChromeDriver to PATH.

workbook = load_workbook(r'C:\Users\Tomas\Desktop\leads sheets\moroccovfair.xlsx')    -> this is the EXACT location of your workbook. It must also be in .xlsx format, not xls

If you want to limit the number of times the script runs, change while x <= 5270:    to the number of times you want the script to run. 
