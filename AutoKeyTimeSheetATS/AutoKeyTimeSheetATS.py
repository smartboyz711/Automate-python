import errno
import os
from datetime import datetime

import pandas as pd
from pandas import ExcelFile
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
defaultTotalhour = 8
df_string = "%d/%m/%Y"

def print_header() :
    print("========================================================================================")
    print("Scirpt auto key time Sheet from excel file for newtimesheet.aware.co.th")
    print("Create By : theedanai Poomilamnao 18/10/2022")
    print("========================================================================================")
    print()

def print_line() :
    print()
    print("========================================================================================")
    print()
class Data_fill :
    
    def __init__(self, customer : str, 
                 project : str, 
                 role : str, 
                 task : str, 
                 billType : str,
                 filldatetime : datetime,
                 hours : float, 
                 description : str,
                 statusMessage : str) -> None:
        
        self.customer = customer
        self.project = project
        self.role = role
        self.task = task
        self.billType = billType
        self.filldatetime = filldatetime
        self.hours = hours
        self.description = description
        self.statusMessage = statusMessage
        
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
            
    def as_dict(self) :
        return {
                    "Datetime": self.filldatetime.strftime(df_string),
                    "Customer" : self.customer,
                    "Project": self.project, 
                    "Role": self.role,
                    "Task": self.task,
                    "BillType": self.billType,
                    "Description": self.description,
                    "Hours": self.hours,
                    "StatusMessage": self.statusMessage
                }
      
def get_driver():
    #set option to make browsing easier
    option = webdriver.ChromeOptions()
    option.add_argument("disable-infobars")
    option.add_argument("start-maximized")
    option.add_argument("disable-dev-shm-usage")
    option.add_argument("no-sandbox")
    option.add_experimental_option(name="excludeSwitches",value=["enable-automation"])
    option.add_argument("disable-blink-features=AutomationControlled")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()) ,options=option)
    driver.get("https://newtimesheet.aware.co.th/timesheet/Login.aspx")
    #driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()))
    return driver
    
def login_timeEntry(driver : WebDriver, username : str, password : str) :
    driver.find_element(By.ID,value="cphContent_txtUserName").send_keys(username)
    driver.find_element(By.ID,value="cphContent_txtUserPassword").send_keys(password + Keys.RETURN)
    #Wait and find Tab Time Entry
    WebDriverWait(driver, timeout=time_out).until(EC.presence_of_element_located((By.LINK_TEXT,"Time Entry")))
    #Click Tab Time Entry 
    driver.find_element(By.LINK_TEXT,value="Time Entry").click() 
    WebDriverWait(driver, timeout=time_out).until(EC.presence_of_element_located((By.ID,"cphContentRight_lblDay"))) 
    driver.find_element(By.ID,value="cphContentRight_lblDay").click()
    WebDriverWait(driver, timeout=time_out).until(EC.presence_of_element_located((By.ID,"cphContent_lblDateShow")))
    
def find_fillDataDate(driver : WebDriver, data_fill : Data_fill) :
    while True :
        #detect date in Calender
        filldatetime = data_fill.filldatetime
        filldate = filldatetime.strftime("%a").upper()
        if(filldate == "SAT" or filldate == "SUN") :
          raise Exception ("Can't fill Time Sheet on Saturday and Sunday")
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
        
        monDateTime = datetime.strptime(f"{dayMon}/{monthMon}/{yearMon}",df_string)
        sunDateTime = datetime.strptime(f"{daySun}/{monthSun}/{yearSun}",df_string)
        
        if(monDateTime <= filldatetime <= sunDateTime) :
            driver.find_element(By.ID,value=filldate).click() #FRI
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
    
    if(float(driver.find_element(By.ID,value="totalHours").text)+data_fill.hours > defaultTotalhour) :
        driver.find_element(By.ID,value="cphContent_DeleteAll").click()
        WebDriverWait(driver, timeout=time_out).until(EC.presence_of_element_located((By.ID,"dialog-confirm-delete")))
        driver.find_element(By.XPATH,value="/html/body/form/div[12]/div[3]/div/button[1]/span").click() #OK
        WebDriverWait(driver, timeout=time_out).until(EC.invisibility_of_element_located((By.ID,"dialog-confirm-delete")))
        WebDriverWait(driver, timeout=time_out).until(EC.invisibility_of_element_located((By.ID,"cphContent_DeleteAll")))

def fill_taskData(driver : WebDriver, data_fill : Data_fill) :
    try :
        driver.find_element(By.ID,value="cphContent_addTimeEntry")
    except NoSuchElementException :
        raise Exception (f"this Datetime {data_fill.filldatetime.strftime('%d/%m/%Y')} is already submitted")

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

def main_fillDataTask(driver : WebDriver, data_fill : Data_fill) -> Data_fill :
    try :
        find_fillDataDate(driver, data_fill)
        delete_allTaskData(driver, data_fill)
        fill_taskData(driver, data_fill)
    except Exception as e:
        data_fill.statusMessage = str(e)
        return data_fill
    data_fill.statusMessage = "Success"
    return data_fill

def convertFileToList(file : ExcelFile) -> list[Data_fill] :
    Data_fill_list = list()
    for sheetname in file.sheet_names:
        datasheet = file.parse(sheetname)
        for i, _ in datasheet.iterrows():
            message = list()
            for column in datasheet.columns:
                match str(column).strip() :
                    case "Customer" : 
                        if(not pd.isnull(datasheet[column][i])) :
                            customer = str(datasheet[column][i]).strip()
                        else :
                            customer = ""
                            message.append("Customer is required field.")
                    case "Project"  :
                        if(not pd.isnull(datasheet[column][i])) :
                            project = str(datasheet[column][i]).strip()
                        else :
                            project = ""
                            message.append("Project is required field.")
                    case "Role" :
                        if(not pd.isnull(datasheet[column][i])) :
                            role = str(datasheet[column][i]).strip()
                        else :
                            role = ""
                            message.append("Role is required field.")
                    case "Task" :
                        if(not pd.isnull(datasheet[column][i])) :
                            task = str(datasheet[column][i]).strip()  
                        else :
                            task = ""
                            message.append("Task is required field.")
                    case "BillType" :
                        if(not pd.isnull(datasheet[column][i])) :
                            billType = str(datasheet[column][i]).strip()  
                        else :
                            billType = ""
                            message.append("billType is required field.")
                    case "Datetime" :
                        filldatetime : datetime
                        if(not pd.isnull(datasheet[column][i])) :
                            try :
                                filldatetime = datasheet[column][i]
                                #filldatetime = datetime.strptime(filldatetime,"%Y-%m-%d %H:%M:%S") 
                            except Exception :
                                message.append("Can't Convert Datetime Please enter format (DD/MM/YYYY) in Datetime format")
                        else :
                            message.append("Datetime is required field.")
                    case "Hours" :
                        hours : float
                        if(not pd.isnull(datasheet[column][i])) :
                            try :
                                hours = float(datasheet[column][i]) 
                            except Exception as e:
                                message.append("Please enter Number for Hours field.")
                            if(hours > defaultTotalhour) :
                                hours = 8
                        else :        
                            message.append("Hours is required field.")
                    case "Description" :
                        if(not pd.isnull(datasheet[column][i])) :
                            description = str(datasheet[column][i]).strip()
                        else :
                            description = ""
                            message.append("Description is required field.")
                            
            if(len(message) > 0) :
                statusMessage = ", ".join(message)
            else :
                statusMessage = ""
                        
            data_fill = Data_fill (
                customer=customer,
                project=project,
                role=role,
                task=task,
                billType=billType,
                filldatetime=filldatetime,
                hours=hours,
                description=description,
                statusMessage=statusMessage
            )
            Data_fill_list.append(data_fill)
    return Data_fill_list

def main() :
    print_header()
    try :
        #Input File Name
        while True :
            fileIn : str = input("Input excel File Name (FileName.xlsx) : ")
            #fileIn = "Time_Sheet.xlsx"
            if(not (fileIn.endswith(".xlsx") or fileIn.endswith(".xls"))) :
                print("FileName is not excel File Please try again.")
                print_line()
                continue
            try:
                file = pd.ExcelFile(fileIn)
            except Exception as e:
                print("Can't read excel file : "+str(e))
                print_line()
                continue
            print()
            #Input User Password
            username : str =  input("Input Username : ")
            #username = ""
            print()
            password : str =  input("Input Password : ")
            #password = ""
            driver = get_driver()
            login_timeEntry(driver, username, password)
            Data_fill_list = convertFileToList(file)
            for i in range(len(Data_fill_list)) :
                if (not Data_fill_list[i].statusMessage) :
                        Data_fill_list[i] = main_fillDataTask(driver,Data_fill_list[i])
                        
            df_data_fill = pd.DataFrame([x.as_dict() for x in Data_fill_list])
            
            #Create Path
            try :
                fileName = fileIn.replace(".xlsx","")
                fileName = fileName.replace(".xls","")
                pathName = fileName+"_Report"
                os.makedirs(pathName)
                outputdir = "{}/{}".format(pathName,pathName+".xlsx")
            except OSError as e:
                # If directory is exists use this directory
                if e.errno == errno.EEXIST:
                    outputdir = "{}/{}".format(pathName,pathName+".xlsx")
                    
            df_data_fill.to_excel(excel_writer=outputdir,
                                  index=False)
            print_line() 
            print("fill time Sheet Success you can check result ==> "+outputdir)
            driver.close()
            break
    except Exception as e :
        print_line() 
        print("An error occurred Cannot Key time sheet. : "+str(e))

if __name__ == "__main__":
    main()
            

        

