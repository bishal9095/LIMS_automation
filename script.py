from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import shutil
import pandas as pd

# Packages to install
# pandas
# openpyxl
# selenium



# Temporary Download Directory(default download directory)
download_dir = "C:\\Users\\bisha\\Downloads"

driver = webdriver.Chrome()

driver.maximize_window()

driver.get("https://livehealth.solutions")

username=driver.find_element(By.NAME,'username')
username.send_keys('pacs-rounak')

password=driver.find_element(By.NAME, 'password')
password.send_keys('Ttly@2241')


signInBtn=driver.find_element(By.XPATH,"//*[@id='lab-form']/div[4]/input")
signInBtn.click()



element_to_hover_over = driver.find_element(By.XPATH,"(//*[@id='hoverDropdown']/a)[3]") 

action_chains = ActionChains(driver)


action_chains.move_to_element(element_to_hover_over).perform()

time.sleep(6)

finance_tab = driver.find_element(By.XPATH,"(//*[@id='hoverDropdown'])[3]/ul/li[3]/a")
finance_tab.click()
time.sleep(8)

MIS_Report=driver.find_element(By.XPATH,"//*[@id='navlink-parent-5']")
MIS_Report.click()
time.sleep(8)


driver.switch_to.window(driver.window_handles[-1])
revenue_report_Tab=driver.find_element(By.XPATH,"//*[@id='doctorReport']")
revenue_report_Tab.click()


select_element = driver.find_element(By.ID,"revenueParameter")
select = Select(select_element)
select.select_by_value("1")  

doctor_input=driver.find_element(By.XPATH,"//*[@id='searchReferralInput']")

# Excel Document file path
file_path = 'C:\\Users\\bisha\\Downloads\\w.xlsx' 
#Sheet Name
sheet_name = 'Sheet1' 
#Column name in the sheet
column_name = 'Doctors'


df = pd.read_excel(file_path, sheet_name=sheet_name)
doctor_names = df[column_name]
for doctor in doctor_names:
    doctor_input.click()
    search_string = str(doctor)
    doctor_input.send_keys(search_string)

    time.sleep(6)

    suggested_doc=driver.find_element(By.XPATH,"//*[@id='doctorReports']/div/div/span/span/div/span/div")
    suggested_doc.click()

    date_picker=driver.find_element(By.XPATH,"//*[@id='misExportDateRange']")
    date_picker.click()
    time.sleep(2)
    last_month=driver.find_element(By.XPATH,"/html/body/div[4]/div[1]/ul/li[6]")
    last_month.click()

    pdf_button=driver.find_element(By.XPATH,"//*[@id='exportPDFDivButton']/div/button[2]")
    pdf_button.click()

    time.sleep(2)
    download_btn=driver.find_element(By.XPATH,"//*[@id='exportPDFDivButton']/div/ul/li[3]/a")
    download_btn.click()

    time.sleep(6)

    downloaded_file = max([download_dir + "\\" + f for f in os.listdir(download_dir)], key=os.path.getctime)
    
    # Download Directory
    new_dir="C:\\Users\\bisha\\Downloads\\perm"
    
    new_file_name = os.path.join(new_dir, f"{search_string}.pdf")  # Set the new file name and extension

    shutil.move(downloaded_file, new_file_name)

    time.sleep(6)

driver.quit()