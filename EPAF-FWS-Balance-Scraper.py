import gspread
from selenium import webdriver
import time
from selenium.webdriver.chrome.service import Service
from datetime import datetime


"""
This script uses Selenium to drive the Chrome browser and iteratively initiate student assistant
hiring PA forms, provided with a CSV file of CWIDs.

All students in the CSV file must have FWS awards; the script will encounter an error if a student
does not have a FWS award.

Your IIT Okta account must be set up with the "security question" 2FA option; email or SMS will not
work for this script.

As PAs for each student are accessed, the original FWS award and the remaining FWS balance will be
extracted from the form's HTML elements and added to the CSV file.

The CSV file cannot be open while the script runs.

This code was adapted from a script that was provided by Rama Sashank Madhurapantula.
"""

"""Set up GSpread Authentication to retrieve a Google Sheet list of student A-Numbers."""
gc = gspread.service_account(filename=r"C:\Users\Sean IIT\PycharmProjects\FWS Balance Tracker\fws-balance-tracker.json")
sh = gc.open("Library Student Assistant FWS Balances Dashboard")
worksheet = sh.sheet1

anumbers = worksheet.col_values(1)
for i in range(1,4):
    anumbers.pop(0)

now = datetime.now()
dt_string = now.strftime("%m/%d/%Y %H:%M:%S")
worksheet.update_acell("B2", str(dt_string))

"""Set up the browser driver."""
service = Service()
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)

"""Provide the URL for the Okta login page and provide login credentials."""
username = input("IIT Username: ")
password = input("IIT Password: ")
sec_question = input("Security Answer: ")



driver.get('https://login.iit.edu/cas/login?service=https%3A%2F%2Fmy.iit.edu%2Fc%2Fportal%2Flogin')
time.sleep(2)
input_uname = driver.find_element("xpath", '//*[@id="input28"]')
input_pswd = driver.find_element("xpath", '//*[@id="input36"]')
input_uname.send_keys(username)
input_pswd.send_keys(password)
btn_login = driver.find_element("xpath", '//*[@id="form20"]/div[2]/input')
btn_login.click()
time.sleep(2)

"""Provide your answer to the Okta security question."""
btn_sec_question = driver.find_element("xpath", '/html/body/div[2]/div[2]/main/div[2]/div/div/div[2]/form/div[2]/div/div[2]/div[2]/div[2]/a')
btn_sec_question.click()
time.sleep(2)

input_security = driver.find_element("xpath", '/html/body/div[2]/div[2]/main/div[2]/div/div/div[2]/form/div[1]/div[4]/div/div[2]/span/input')
input_security.send_keys(sec_question)
btn_seclogin = driver.find_element("xpath", '/html/body/div[2]/div[2]/main/div[2]/div/div/div[2]/form/div[2]/input')
btn_seclogin.click()
time.sleep(5)

"""Open a Chrome browser instance."""
driver.execute_script("window.open('');")
driver.switch_to.window(driver.window_handles[1])




for i in anumbers:
    driver.get('https://wildfly-prd.iit.edu/StudentEPAF/')
    input_studentId = driver.find_element("id", 'studentId')
    input_org = driver.find_element("id", 'organization')
    btn_newPA = driver.find_element("id", 'go')

    input_studentId.clear()
    input_org.clear()
    input_studentId.send_keys(i)
    input_org.send_keys("2011")  # INSERT YOUR ORG CODE HERE
    btn_newPA.click()
    time.sleep(2)
    driver.switch_to.window(driver.window_handles[2])

    try:
        FWS_Amount_Element = driver.find_element("xpath", '/html/body/div[2]/div/div/div/div/div/div/div/section/div/div/div/div/div/div/form/table/tbody/tr[2]/td/table/tbody/tr[6]/td[4]/label')
        FWS_Amount = FWS_Amount_Element.text

    except:
        FWS_Amount_Element = "NO FWS"
        FWS_Amount = FWS_Amount_Element


    try:
        FWS_Balance_Element = driver.find_element("xpath", '/html/body/div[2]/div/div/div/div/div/div/div/section/div/div/div/div/div/div/form/table/tbody/tr[2]/td/table/tbody/tr[6]/td[6]/label')
        FWS_Balance = FWS_Balance_Element.text

    except:
        FWS_Balance_Element = "NO FWS"
        FWS_Balance = FWS_Balance_Element

    print(FWS_Amount + ", " + FWS_Balance)
    time.sleep(3)

    worksheet.update_cell(anumbers.index(i) + 4, 5, FWS_Amount)
    worksheet.update_cell(anumbers.index(i) + 4, 6, FWS_Balance)

    driver.close()
    time.sleep(2)
    driver.switch_to.window(driver.window_handles[1])
