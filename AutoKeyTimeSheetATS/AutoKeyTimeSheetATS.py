import time
from datetime import datetime

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

time_out = 10

class Data_fill :
    
    def __init__(self, customer : str, 
                 project : str, 
                 role : str, 
                 task : str, 
                 billType : str,
                 filldatetime : datetime,
                 hours : float, 
                 description : str) :
        
        self.customer = customer
        self.project = project
        self.role = role
        self.task = task
        self.billType = billType
        self.filldatetime = filldatetime
        self.hours = hours
        self.description = description
        
    def get_id_billtype(self) :
        match self.billType :
            case "Regular" :
                return "cphContent_rdoPopBillType_0"
            case "Overtime" :
                return "cphContent_rdoPopBillType_1"
            case "Non-Billable" :
                return "cphContent_rdoPopBillType_2"
            case "Overtime Nonbill" :
                return "cphContent_rdoPopBillType_3"
            case _:
                return ""
   
def get_driver():
    #set option to make browsing easier
    url = "https://newtimesheet.aware.co.th/timesheet/Login.aspx"
    option = webdriver.ChromeOptions()
    option.add_argument("disable-infobars")
    option.add_argument("start-maximized")
    option.add_argument("disable-dev-shm-usage")
    option.add_argument("no-sandbox")
    option.add_experimental_option(name="excludeSwitches",value=["enable-automation"])
    option.add_argument("disable-blink-features=AutomationControlled")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()) ,options=option)
    driver.get(url)
    #driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()))
    return driver
    
def login_timeEntry(driver : WebDriver, username : str, password : str) :
    driver.find_element(By.ID,value="cphContent_txtUserName").send_keys(username)
    time.sleep(1)
    driver.find_element(By.ID,value="cphContent_txtUserPassword").send_keys(password + Keys.RETURN)
    #Wait and find Tab Time Entry
    WebDriverWait(driver, timeout=time_out).until(EC.presence_of_element_located((By.LINK_TEXT,"Time Entry")))
    #Click Tab Time Entry 
    driver.find_element(By.LINK_TEXT,value="Time Entry").click() 
    WebDriverWait(driver, timeout=time_out).until(EC.presence_of_element_located((By.ID,"cphContentRight_lblDay"))) 
    driver.find_element(By.ID,value="cphContentRight_lblDay").click()
    WebDriverWait(driver, timeout=time_out).until(EC.presence_of_element_located((By.ID,"cphContent_lblDateShow")))
    
def find_fillDataDate(driver : WebDriver, filldatetime : datetime) :
    while True :
        #detect date in Calender
        #get day Monday
        elementMon = driver.find_element(By.ID,value="MON")
        dayMon = elementMon.find_element(By.CLASS_NAME,value="day-Num").get_attribute("textContent")
        monthMon = str(int(elementMon.find_element(By.CLASS_NAME,value="month-Num").get_attribute("textContent"))+1)
        yearMon = elementMon.find_element(By.CLASS_NAME,value="year-Num").get_attribute("textContent")
        #get day Sunday
        elementSun = driver.find_element(By.ID,value="SUN")
        daySun = elementSun.find_element(By.CLASS_NAME,value="day-Num").get_attribute("textContent")
        monthSun =  str(int(elementSun.find_element(By.CLASS_NAME,value="month-Num").get_attribute("textContent"))+1)
        yearSun = elementSun.find_element(By.CLASS_NAME,value="year-Num").get_attribute("textContent")
        
        df_string = "%d/%m/%Y"
        monDateTime = datetime.strptime(f"{dayMon}/{monthMon}/{yearMon}",df_string)
        sunDateTime = datetime.strptime(f"{daySun}/{monthSun}/{yearSun}",df_string)
        
        if(monDateTime <= filldatetime <= sunDateTime) :
            driver.find_element(By.ID,value=filldatetime.strftime("%a").upper()).click() #FRI
            WebDriverWait(driver, timeout=time_out).until(EC.text_to_be_present_in_element
                                                            ((By.ID,"cphContent_lblDateShow"),
                                                            filldatetime.strftime("%A, %B %d, %Y"))) #Friday, October 14, 2022
            break
        elif(filldatetime < monDateTime) :
            driver.find_element(By.CLASS_NAME,value="previousWeek").click()
            WebDriverWait(driver, timeout=time_out).until(EC.element_to_be_clickable((By.CLASS_NAME,"previousWeek")))
            continue
        else :
            driver.find_element(By.CLASS_NAME,value="nextWeek").click()
            WebDriverWait(driver, timeout=time_out).until(EC.element_to_be_clickable((By.CLASS_NAME,"nextWeek")))
            continue

def delete_allTaskData(driver : WebDriver, data_fill : Data_fill) :
    try :
        driver.find_element(By.ID,value="cphContent_DeleteAll")
    except NoSuchElementException :
        return
    if(float(driver.find_element(By.ID,value="totalHours").text)+data_fill.hours > 8) :
        driver.find_element(By.ID,value="cphContent_DeleteAll").click()
        WebDriverWait(driver, timeout=time_out).until(EC.presence_of_element_located((By.ID,"dialog-confirm-delete")))
        driver.find_element(By.XPATH,value="/html/body/form/div[12]/div[3]/div/button[1]/span").click() #OK
        WebDriverWait(driver, timeout=time_out).until(EC.invisibility_of_element_located((By.ID,"dialog-confirm-delete")))
        WebDriverWait(driver, timeout=time_out).until(EC.invisibility_of_element_located((By.ID,"cphContent_DeleteAll")))

def fill_TaskData(driver : WebDriver, data_fill : Data_fill) :
    try :
        driver.find_element(By.ID,value="cphContent_addTimeEntry")
    except NoSuchElementException :
        return f"this Datetime {data_fill.filldatetime.strftime('%d/%m/%Y')} is already submitted"
    
    try :
        driver.find_element(By.ID,value="cphContent_addTimeEntry").click()
        WebDriverWait(driver, timeout=time_out).until(EC.presence_of_element_located((By.ID,"cphContent_pnlAddEditTimelist")))
        pnlAddEditTimelist = driver.find_element(By.ID,value="cphContent_pnlAddEditTimelist")
        driver.find_element(By.ID,value="cphContent_lnkAddTimelist").click()
        WebDriverWait(driver, timeout=time_out).until(EC.text_to_be_present_in_element((By.ID,"cphContent_ddlPopCustomer"),data_fill.customer))
        Select(pnlAddEditTimelist.find_element(By.ID,value="cphContent_ddlPopCustomer")).select_by_visible_text(data_fill.customer)
        WebDriverWait(driver, timeout=time_out).until(EC.text_to_be_present_in_element((By.ID,"cphContent_ddlPopProject"),data_fill.project))
        Select(pnlAddEditTimelist.find_element(By.ID,value="cphContent_ddlPopProject")).select_by_visible_text(data_fill.project)
        WebDriverWait(driver, timeout=time_out).until(EC.text_to_be_present_in_element((By.ID,"cphContent_ddlPopRole"),data_fill.role))
        Select(pnlAddEditTimelist.find_element(By.ID,value="cphContent_ddlPopRole")).select_by_visible_text(data_fill.role)
        WebDriverWait(driver, timeout=time_out).until(EC.text_to_be_present_in_element((By.ID,"cphContent_ddlPopTask"),data_fill.task))
        Select(pnlAddEditTimelist.find_element(By.ID,value="cphContent_ddlPopTask")).select_by_visible_text(data_fill.task)
        pnlAddEditTimelist.find_element(By.ID,value=data_fill.get_id_billtype()).click()
        pnlAddEditTimelist.find_element(By.ID,value="cphContent_txtHours").send_keys(data_fill.hours)
        pnlAddEditTimelist.find_element(By.ID,value="cphContent_rdlInternal1").click()
        Select(pnlAddEditTimelist.find_element(By.ID,value="cphContent_ddlInternalDescription")).select_by_visible_text(data_fill.task)
        pnlAddEditTimelist.find_element(By.ID,value="cphContent_rdlInternal2").click()
        pnlAddEditTimelist.find_element(By.ID,value="cphContent_txtInternalDescription").send_keys(data_fill.description)
        driver.find_element(By.XPATH,value="/html/body/form/div[12]/div[11]/div/button[1]/span").click() #save
        WebDriverWait(driver, timeout=time_out).until(EC.invisibility_of_element_located((By.ID,"cphContent_pnlAddEditTimelist")))
        return "Success"
    except Exception as e :
        return str(e)
        
if __name__ == "__main__":
    try :
        driver = get_driver()
        username = ""
        password = ""
        data_fill = Data_fill(
            customer="AIS",
            project="AIS-SFF",
            role="Programmer",
            task="Coding : Java",
            billType="Regular",
            filldatetime=datetime.strptime("16/11/2022","%d/%m/%Y"),
            hours=3.5,
            description="Test fill time Sheet"
        )
        login_timeEntry(driver, username, password)
        find_fillDataDate(driver, data_fill.filldatetime)
        delete_allTaskData(driver, data_fill)
        print(fill_TaskData(driver, data_fill))
    except Exception as e :
        print("An exception occurred Cannot key time sheet. : "+str(e))

