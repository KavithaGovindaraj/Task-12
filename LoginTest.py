from Locators.Homepage import OrangeHRM_Locators
from Utilities.excel_functions import SumanExcelFunctions
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager


excel_file = OrangeHRM_Locators().excel_file


sheet_number = OrangeHRM_Locators().sheet_number


# create object for the Excel Utility Class
suman = SumanExcelFunctions(excel_file, sheet_number)


driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.get(OrangeHRM_Locators().url)
driver.implicitly_wait(5)


# row count from the Excel file
row = suman.row_count()


for row in range(2, row+1):
   username = suman.read_data(row, 6)
   password = suman.read_data(row, 7)


   driver.find_element(by=By.NAME, value=OrangeHRM_Locators().username).send_keys(username)
   driver.find_element(by=By.NAME, value=OrangeHRM_Locators().password).send_keys(password)
   driver.find_element(by=By.XPATH, value=OrangeHRM_Locators().submit_button).click()


   driver.implicitly_wait(5)

# validate the login and generate the Test-Case results & reports
   if OrangeHRM_Locators().dashboard_url == driver.current_url:
       print("SUCCESS : Login with Username {a} & Password {b}".format(a=username, b=password))
       suman.write_data(row, 8, OrangeHRM_Locators().pass_data)
       driver.back()
   elif OrangeHRM_Locators().url == driver.current_url:
       print("ERROR : Login unsuccessful with username {a} & Password {b}".format(a=username, b=password))
       suman.write_data(row, 8, OrangeHRM_Locators().fail_data)
       driver.refresh()


driver.quit()

# Output:

#C:\Users\PREMA\PAT-28\DDTF1\venv\Scripts\python.exe C:\Users\PREMA\PAT-28\DDTF1\Testcases\LoginTest.py
#SUCCESS : Login with Username Admin & Password admin123
#ERROR : Login unsuccessful with username ramesh & Password ramesh@123
#Traceback (most recent call last):
#  File "C:\Users\PREMA\PAT-28\DDTF1\Testcases\LoginTest.py", line 33, in <module>
#    driver.find_element(by=By.NAME, value=OrangeHRM_Locators().username).send_keys(username)
#    ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
#  File "C:\Users\PREMA\PAT-28\DDTF1\venv\Lib\site-packages\selenium\webdriver\remote\webdriver.py", line 748, in find_element
#    return self.execute(Command.FIND_ELEMENT, {"using": by, "value": value})["value"]
#           ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
#  File "C:\Users\PREMA\PAT-28\DDTF1\venv\Lib\site-packages\selenium\webdriver\remote\webdriver.py", line 354, in execute
#    self.error_handler.check_response(response)
#  File "C:\Users\PREMA\PAT-28\DDTF1\venv\Lib\site-packages\selenium\webdriver\remote\errorhandler.py", line 229, in check_response
#    raise exception_class(message, screen, stacktrace)
#selenium.common.exceptions.NoSuchElementException: Message: no such element: Unable to locate element: {"method":"css selector","selector":"[name="username"]"}
#  (Session info: chrome=126.0.6478.185); For documentation on this error, please visit: https://www.selenium.dev/documentation/webdriver/troubleshooting/errors#no-such-element-exception
#Stacktrace:
#	GetHandleVerifier [0x006CC203+27395]
#	(No symbol) [0x00663E04]
#	(No symbol) [0x00561B7F]
#	(No symbol) [0x005A2C65]
#	(No symbol) [0x005A2D3B]
#	(No symbol) [0x005DEC82]
#	(No symbol) [0x005C39E4]
#	(No symbol) [0x005DCB24]
#	(No symbol) [0x005C3736]
#	(No symbol) [0x00597541]
#	(No symbol) [0x005980BD]
#	GetHandleVerifier [0x00983AB3+2876339]
#	GetHandleVerifier [0x009D7F7D+3221629]
#	GetHandleVerifier [0x0074D674+556916]
#	GetHandleVerifier [0x0075478C+585868]
#	(No symbol) [0x0066CE44]
#	(No symbol) [0x00669858]
#	(No symbol) [0x006699F7]
#	(No symbol) [0x0065BF4E]
#	BaseThreadInitThunk [0x751C7BA9+25]
#	RtlInitializeExceptionChain [0x7746C10B+107]
#	RtlClearBits [0x7746C08F+191]
#
#
#Process finished with exit code 1
