from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime
import os
import time
import requests
EDGE_DRIVER_PATH="C:\\Users\\Downloads\\edgedriver_win64\\msedgedriver.exe"
EDGE_PROFILE_PATH="C:\\Users\\AppData\\Local\\Microsoft\\Edge\\User Data\\Work"
edge_options=Options()
edge_options.use_chromium=True
edge_options.set_capability("goog:loggingPrefs",{"performance":"ALL"})
edge_options.add_argument("--disable-headless-mode")
edge_options.add_argument(f"--user-data-dir={EDGE_PROFILE_PATH}")
edge_options.add_argument("--profile-directory=Default")
edge_options.add_argument("--start-maximized")
service=Service(EDGE_DRIVER_PATH)
driver=webdriver.Edge(service=service,options=edge_options)
output_file="workdayreport2.xls"
time.sleep(20)

def sharepointupload(file_to_upload):
    driver.get("Sharepoint Location")
    #time.sleep(5)
    update_button=WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, "//button[contains(@aria-label,'Upload')]"))
    )
    ActionChains(driver).move_to_element(update_button).click().perform()
    time.sleep(2)
    file_Action=WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, "//span[text()='Files']"))
    )
    ActionChains(driver).move_to_element(file_Action).click().perform()
    time.sleep(2)
    file_input=WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
    )
    file_input.send_keys(file_to_upload)
    time.sleep(20)
    driver.quit()
#hqve to run it in admin mode
#cookies=browser_cookie3.edge(domain_name="myworkday.com")
try:
    driver.get("https://www.myworkday.com/{domain}")
    print("identification in progress")
    print("Started Identifying")
    profile_icon=WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="app-chrome-container"]/div/div[6]/div[3]/div[4]'))
    )
    ActionChains(driver).move_to_element(profile_icon).click().perform()
    viewReport=WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="app-chrome-container"]/div/div[6]/div[3]/div[4]/div[2]/ul/div/li[4]'))
    )
    ActionChains(driver).move_to_element(viewReport).click().perform()
    time.sleep(10)
    element=driver.find_element(By.XPATH, '//*[@id="riva-grid-api-key-uid4"]/div/div/div/div[1]/div/table/tbody/tr[1]/td[1]')
    actions=ActionChains(driver)
    actions.context_click(element).perform()
    time.sleep(5)
    copyUrl=WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "div[data-automation-id='copyUrl']"))
    )
    print("element found")
    download_url=copyUrl.get_attribute("data-clipboard-text")
    print("downlaod--------",download_url)
     #using local storage cookies
  #  cookies={
   #     "JSESSIONID":"",
    #    "wd-browser-id":"",
     #   "wd-alt-sessionid":""
   # }
    session=requests.Session()
    for cookie in driver.get_cookies():
        session.cookies.set(cookie['name'],cookie['value'])
    first_response=session.get(download_url)
    print("rtexr",first_response.text)
    print("",first_response)
    with open(output_file,"wb") as f:
        f.write(first_response.content)
    print("file saved")
    time.sleep(5)
    filepath="C:\PythonCode\{output_file}"
    if os.path.exists(filepath):
        print("File downloaded successfully")
       #sharepointupload(filepath)
    else:
        print("error in downloading the file")
    
  
except Exception as e:
    print(f"Error clicking on the button: {e}")  
    print("Report Downloaded")
    time.sleep(5)
finally:
    driver.quit()

print("Execution Successfull")
