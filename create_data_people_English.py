from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import random
import logging
from response import catch_response 
from selenium.webdriver.support.ui import Select
import string

def source_input_English(driver,url):
    time.sleep(5)
    
    source_type = ['Mobile Cat.1','Stationary Cat.1','Processes Cat.1','Fugitive Cat.1','Others GHG Cat.1','Electricity Cat.2','Visitor Cat.3','Business travel Cat.3','Employee commuting Cat.3','Waste Cat.4',
                   'Capital goods Cat.4','US Oil & electricity Cat.4','Water B.5.2.a','Purchased goods Cat.4','Leased assets Cat.4','Services Cat.4','US transport Cat.3','Waste transport Cat.3',
                   'DS transport Cat.3','DS leased assets Cat.5','Product used Cat.5','Investments Cat.5','EOL Cat.5', 'Other emissions Cat.6']
    url = url+"/calculate/" 

    for i in source_type:
        if i=='Mobile Cat.1' or i == 'Stationary Cat.1'or i == 'Processes Cat.1' or i== 'Others GHG Cat.1':
            link = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.LINK_TEXT, i)))
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))  
            link.click()
            for j in range(3):
                element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), 'Add')]")))
                element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Add')]")))
                driver.execute_script("arguments[0].click()", element)
                time.sleep(0.5)
                input_data_C1_C2(driver, i,j,0)
                # input_data_C1_C2_verify(driver)
                
        
            
        elif i == 'Fugitive Cat.1':
            link = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.LINK_TEXT, i)))
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))  
            link.click()
            time.sleep(0.5)
            for j in range(1):
                workhour(driver,url,j)
                time.sleep(0.5)
                
            element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[text()="Refrigeration equipment B.2.2.D"]')))
            element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="Refrigeration equipment B.2.2.D"]')))
            driver.execute_script("arguments[0].click()", element)
            for j in range(3):
                element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), 'Add')]"))) 
                element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Add')]")))
                driver.execute_script("arguments[0].click()", element)
                time.sleep(0.5)
                input_data_C1_C2(driver, i,j,1)
                time.sleep(0.5)

            element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[text()="Firefighting equipment B.2.2.D"]')))
            element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="Firefighting equipment B.2.2.D"]')))
            driver.execute_script("arguments[0].click()", element)
            for j in range(3):
                element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), 'Add')]")))
                element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Add')]")))
                driver.execute_script("arguments[0].click()", element)
                time.sleep(0.5)
                input_data_C1_C2(driver, i,j,2)
                time.sleep(0.5)
        elif i == 'US Oil & electricity Cat.4':
            link = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.LINK_TEXT, i)))
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))  
            link.click()

            for j  in range(3):              
                element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), 'Add')]")))
                element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Add')]")))
                driver.execute_script("arguments[0].click()", element)    
                time.sleep(0.5)
                input_data_C3_C6(driver, i,j,3) 
                time.sleep(0.5)
            
            element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[text()="Upstream emissions of other energy imported B.5.2.a"]')))
            element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="Upstream emissions of other energy imported B.5.2.a"]')))
            driver.execute_script("arguments[0].click()", element)
            for j in range(3):
                element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), 'Add')]")))
                element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Add')]")))
                driver.execute_script("arguments[0].click()", element)
                time.sleep(0.5)
                input_data_C3_C6(driver, i,j,4)
        elif i == 'Electricity Cat.2':
            link = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.LINK_TEXT, i)))
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))  
            link.click()
            for j in range(1):
                elec(driver,url,j)
                time.sleep(0.5)
            #     driver.get(url+"electricity")
            #     time.sleep(0.5)
                
            # time.sleep(0.5)
            # green_elec(driver,url)
        # elif i == '輸入蒸汽 C2' :
        #     link = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.LINK_TEXT, i)))
        #     link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))  
        #     link.click()
        #     for j in range(1):
        #         time.sleep(1)
        #         steam(driver,url,j)
        #         time.sleep(0.5)
                
        else:
            link = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.LINK_TEXT, i)))
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))  
            link.click()
            
            for j in range(3):
                element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), 'Add')]")))
                element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Add')]")))
                driver.execute_script("arguments[0].click()", element)
                time.sleep(0.5)
                input_data_C3_C6(driver, i,j,0)
                time.sleep(0.5)
        
        driver.get(url)
    # driver.quit()
def wait_and_click(driver, by, value, wait_time=2):
    element = WebDriverWait(driver, wait_time).until(EC.element_to_be_clickable((by, value)))
    element.click()
    time.sleep(1)  # 若需要等待一段時間
    
def coefficient_C4_C6(driver):
    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "validateOnly")))
    time.sleep(2)
    wait_and_click(driver, By.XPATH, "//button[@class='ant-btn css-dts6b9 ant-btn-link d-flex align-items-center mt-2']")
    
    wait_and_click(driver, By.XPATH, "//div[@class='ant-collapse-item'][2]")
    
    wait_and_click(driver, By.XPATH, "//div[@class='ant-collapse-item']//span[contains(text(), '01農、林、漁業及相關活動')]")
    
    wait_and_click(driver, By.XPATH, "//a[@class='d-block my-2 text-dark'][1]")

    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//tr[@class='ant-table-row ant-table-row-level-0'][1]//span[@class='ant-checkbox-inner']")))
    actions = ActionChains(driver)
    actions.move_to_element(element).click().perform()
    time.sleep(1)
    
    wait_and_click(driver, By.XPATH, "//button[@class='ant-btn css-dts6b9 ant-btn-primary']//span[contains(text(),'OK')]")

def coefficient_C3(driver):
    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "validateOnly")))
    time.sleep(2)
    wait_and_click(driver, By.XPATH, "//button[@class='ant-btn css-dts6b9 ant-btn-link d-flex align-items-center mt-2']")
    
    wait_and_click(driver, By.XPATH, "//div[@class='ant-collapse-item'][2]")
    
    wait_and_click(driver, By.XPATH, "//div[@class='ant-collapse-item']//span[contains(text(), '25 shipment service')]")
    
    wait_and_click(driver, By.XPATH, "//a[@class='d-block my-2 text-dark'][1]")

    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//tr[@class='ant-table-row ant-table-row-level-0'][1]//span[@class='ant-checkbox-inner']")))
    actions = ActionChains(driver)
    actions.move_to_element(element).click().perform()
    time.sleep(1)
    
    wait_and_click(driver, By.XPATH, "//button[@class='ant-btn css-dts6b9 ant-btn-primary']//span[contains(text(),'OK')]")
def coefficient_C4(driver):
    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "validateOnly")))
    time.sleep(2)
    wait_and_click(driver, By.XPATH, "//button[@class='ant-btn css-dts6b9 ant-btn-link d-flex align-items-center mt-2']")
    
    wait_and_click(driver, By.XPATH, "//div[@class='ant-collapse-item'][2]")
    
    wait_and_click(driver, By.XPATH, "//div[@class='ant-collapse-item']//span[contains(text(), '26 waste recycling services')]")
    
    wait_and_click(driver, By.XPATH, "//a[@class='d-block my-2 text-dark'][1]")

    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//tr[@class='ant-table-row ant-table-row-level-0'][1]//span[@class='ant-checkbox-inner']")))
    actions = ActionChains(driver)
    actions.move_to_element(element).click().perform()
    time.sleep(1)
    
    wait_and_click(driver, By.XPATH, "//button[@class='ant-btn css-dts6b9 ant-btn-primary']//span[contains(text(),'OK')]")


def input_data_C3_C6(driver, source_type,index,rfg_or_fire):
    # random_letters = ''.join(random.choices(string.ascii_letters, k=10))
    validateOnlyCommuting = ["Travelling on other modes of transport  ","Car-Petrol ","Car-Diesel "]
    validateOnlyGHGEvaluateItem = ["Assessment Project Test 1","Assessment Project Test 2","Assessment Project Test 3","Assessment Project Test 4","Assessment Project Test 5"]
    validateOnlyMaterialNo = ["Material number test 1","Material number test 2","Material number test 3","Material number test 4","Material number test 5"]
    validateOnlyMaterialName = ["Material name test 1","Material name test 2","Material name test 3","Material name test 4","Material name test 5"]
    validateOnlyCargoName = ["Assessment Project Test 1","Assessment Project Test 2"," Assessment Project Test 3","Assessment Project Test 4","Assessment Project Test 5"]
    validateOnlyTransportTypeID = ["Road,Transport lorry 3.5-7.5t EURO5/RER S","Road,Transport lorry 7.5-16t EURO5/RER S ","Road,Light commercial vehicle","Road,Lorry with cooling 7.5-16 ton","Road,Lorry with freezing 7.5-16 ton"]
    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "validateOnly")))
    source_types_to_check = {
                            "Visitor Cat.3", 
                            "Business travel Cat.3", 
                            "Employee commuting Cat.3", 
                            "US transport Cat.3", 
                            "Waste transport Cat.3", 
                            "DS transport Cat.3"
                            }
    if source_type in source_types_to_check:
        coefficient_C3(driver)
        time.sleep(1)
    elif source_type == 'Waste Cat.4':
        coefficient_C4(driver)
        time.sleep(1)
    else:
        coefficient_C4_C6(driver)
        time.sleep(1)
    

    form_element = driver.find_element(By.ID, "validateOnly")               #表單element

    input_elements = form_element.find_elements(By.TAG_NAME, "input")       #tag_name是input的
    
    new_inputelements =[]

    for i in input_elements:                     #把disable的先移除掉
        if i.get_attribute("disabled") != None and i.get_attribute("id")!="validateOnly_referenceFile" and i.get_attribute("type")!="search":
            print("")
        else:
            new_inputelements.append(i)

    time.sleep(0.5)
    for input_element in new_inputelements:
        try:
            if input_element.get_attribute("placeholder") == "Please enter" and input_element.get_attribute("id") == 'validateOnly_GHGEvaluateItem': #評估項目
                if "Waste" in source_type:
                    prefix = "Waste "
                elif "Capital goods" in source_type:
                    prefix = "Capital goods "
                elif "Water" in source_type:
                    prefix = "Water "
                elif "Purchased goods" in source_type:
                    prefix = "Purchased goods"
                elif "Leased assets" in source_type:
                    prefix = "Leased assets "
                elif "DS leased assets" in source_type:
                    prefix = "DS leased assets "
                elif "Product used " in source_type:
                    prefix = "Product used "
                elif "Investments" in source_type:
                    prefix = "Investments "
                elif "EOL" in source_type:
                    prefix = "EOL "
                else:
                    prefix = ""
                
                validateOnly_GHGEvaluateItem = prefix + validateOnlyGHGEvaluateItem[index % len(validateOnlyGHGEvaluateItem)]
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(validateOnly_GHGEvaluateItem)
            elif input_element.get_attribute("placeholder") == "Please enter" and input_element.get_attribute("id") == 'validateOnly_MaterialNo': #料號
                validateOnly_MaterialNo = validateOnlyMaterialNo[index % len(validateOnlyMaterialNo)]
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(validateOnly_MaterialNo)
            elif input_element.get_attribute("placeholder") == "Please enter" and input_element.get_attribute("id") == 'validateOnly_ActivityIntensityUnit': #活動強度(使用量)單位
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("Kg")
                
            elif input_element.get_attribute("placeholder") == "Please enter" and input_element.get_attribute("id") == 'validateOnly_MaterialName': #物料名稱
                validateOnly_MaterialName = validateOnlyMaterialName[index % len(validateOnlyMaterialName)]
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(validateOnly_MaterialName)
                
            elif input_element.get_attribute("placeholder") == "Please enter" and input_element.get_attribute("id") == 'validateOnly_CargoName': #上游評估項目
                validateOnly_CargoName = validateOnlyCargoName[index % len(validateOnlyCargoName)]
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(validateOnly_CargoName)
                
            
            elif input_element.get_attribute("id") == 'validateOnly_ingredientName':
                continue

            elif input_element.get_attribute("placeholder") == "Please enter the number" and input_element.get_attribute("id") == 'validateOnly_ActivityIntensity': #活動強度(使用量)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(100)
                
            elif input_element.get_attribute("placeholder") == "Please enter the number" and input_element.get_attribute("id") == 'validateOnly_AnnualPurchaseAmount': #年度採購量
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(100)
                
            elif input_element.get_attribute("placeholder") == "Please enter the number" and input_element.get_attribute("id") == 'validateOnly_TransportWeight': #上游重量
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(10)
                
            elif input_element.get_attribute("placeholder") == "Please enter the number" and input_element.get_attribute("id") == 'validateOnly_TransportDistance': #上游運輸距離(km)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(10)
                
            elif input_element.get_attribute("placeholder") == "Please enter the number" and input_element.get_attribute("id") == 'validateOnly_TotalEmployees':#同行人數
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(10)
                
            elif input_element.get_attribute("placeholder") == "Please enter the number" and input_element.get_attribute("id") == 'validateOnly_WorkingDays': #盤查期間次數
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(24)
                
            elif input_element.get_attribute("placeholder") == "Please enter the number" and input_element.get_attribute("id") == 'validateOnly_AverageMovingDistance': #移動距離
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(100)
                
            elif input_element.get_attribute("placeholder") == "Please enter the number" and input_element.get_attribute("id") == 'validateOnly_MaterialSpec': #單一物料重量
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(100)
                
                
            elif input_element.get_attribute("placeholder") == "Please enter" and input_element.get_attribute("id") == 'validateOnly_ActivityIntensityUnit': #單一物料重量單位
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("Kg")
                
   
            elif input_element.get_attribute("type")=="search":   #選單類 先click 再選 再click
        
                if input_element.get_attribute("id")=="validateOnly_activityDataType":            
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Automatic Continuous Measurement')]")))
                    element.click()
                
                elif input_element.get_attribute("id")=="validateOnly_emitParaType":
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Mass Energy Balance')]")))
                    element.click()
                
                elif input_element.get_attribute("id") == "validateOnly_DataSource":
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Global or regional level factor')]")))
                    element.click()
                
                elif input_element.get_attribute("id") == "validateOnly_DataAttribute":
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Based on a hypothetical situation')]")))
                    element.click()
                
                elif input_element.get_attribute("id") == "validateOnly_Commuting": #員工差旅交通方式
                    input_element.click()
                    validateOnly_Commuting = validateOnlyCommuting[index % len(validateOnlyCommuting)]
                    element1 = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, f"//div[contains(text(), '{validateOnly_Commuting}')]")))
                    element1.click()     
                
                elif input_element.get_attribute("id") == "validateOnly_TransportTypeID": #運輸方式
                    input_element.click()
                    validateOnly_TransportTypeID = validateOnlyTransportTypeID[index % len(validateOnlyTransportTypeID)]
                    element1 = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, f"//div[contains(text(), '{validateOnly_TransportTypeID}')]")))
                    element1.click()      
                
                else:
                    continue  
                
            
        except : #StaleElementReferenceException
            print("")
            
                    
    button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']")))    #送出
    button.click()
    time.sleep(2)
    if rfg_or_fire == 1:
        catch_response(driver,'Fugitive Cat.19(RFG)')
    elif rfg_or_fire == 2:
        catch_response(driver,'Fugitive Cat.1(Fire)')
    elif rfg_or_fire == 3:
        catch_response(driver, 'US Oil & electricity Cat.4(Upstream emissions of electricity imported)')
    elif rfg_or_fire == 4:
        catch_response(driver, 'US Oil & electricity Cat.4(Upstream emissions of other energy imported)')
    elif rfg_or_fire == 0:
        catch_response(driver,source_type)


def input_data_C1_C2_verify(driver):
    try:
        table = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, "//div[@class='ant-table-content']")))
        print(table.text)
    except:
        print("沒有")
    
    
    


def input_data_C1_C2(driver, source_type,index,rfg_or_fire):
    # random_letters = ''.join(random.choices(string.ascii_letters, k=10))
    
    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "validateOnly")))
    
    form_element = driver.find_element(By.ID, "validateOnly")               #表單element

    input_elements = form_element.find_elements(By.TAG_NAME, "input")       #tag_name是input的
    CarModel = ["Business vehicle test 1","Business vehicle test 2","Business vehicle test 3"]
    CarPlateNo = ["ABC-123","ABC-234","ABC-345","ABC-456","ABC-567"]
    validateOnlyName1 = ["Fixed device name test 1","Fixed device name test 2","Fixed device name test 3","Fixed device name test 4","Fixed device name test 5"]
    validateOnlyName2 = ["Process equipment name test 1","Process equipment name test 2","Process equipment name test 3","Process equipment name test 4","Process equipment name test 5"]
    validateOnlyName3 = ["Device name test 1","Device name test 2","Device name test 3","Device name test 4","Device name test 5"]
    Area = ["台灣","大陸","越南","泰國","墨西哥"]
    validateOnlyParameterID = ["R-11","R-12","R-13","R-112","R-112a"]
    validateOnlyParameterID1 = ["R-123","R-125","FM-200","HFC-236fa","Halon-1201"]
    validateOnlyParameterID2 = ["Household refrigerator","Water dispenser","Dehumidifier","Refrigerator","Constant temperature and humidity"]
    
    
    new_inputelements =[]

    for i in input_elements:                     #把disable的先移除掉
        if i.get_attribute("disabled") != None and i.get_attribute("id")!="validateOnly_referenceFile" and i.get_attribute("type")!="search":
            print("")
        else:
            new_inputelements.append(i)
    for input_element in new_inputelements:
        try:
            if input_element.get_attribute("placeholder") == "Please enter" and input_element.get_attribute("id") == 'validateOnly_CarPlateNo': #車牌
                car_plate = CarPlateNo[index % len(CarPlateNo)]
                WebDriverWait(driver, 10).until(EC.visibility_of(input_element)).send_keys(str(car_plate))
                
            elif input_element.get_attribute("placeholder") == "Please enter" and input_element.get_attribute("id") == 'validateOnly_ProcessName': #製程別
                WebDriverWait(driver, 10).until(EC.visibility_of(input_element)).send_keys("Process specific testing")
                
            elif input_element.get_attribute("placeholder") == "Please enter" and input_element.get_attribute("id") == 'validateOnly_ResponsibleUnit': #負責單位
                WebDriverWait(driver, 10).until(EC.visibility_of(input_element)).send_keys("Responsible for unit testing")
                
            elif input_element.get_attribute("placeholder") == "Please enter" and input_element.get_attribute("id") == 'validateOnly_Name': #公務車型/設備名稱
                if source_type == "Mobile Cat.1":
                    car_model = CarModel[index % len(CarModel)]
                    WebDriverWait(driver, 10).until(EC.visibility_of(input_element)).send_keys(str(car_model))
                    
                    time.sleep(0.5)
                elif source_type == "Stationary Cat.1":
                    validateonlyname1 = validateOnlyName1[index % len(validateOnlyName1)]
                    WebDriverWait(driver, 10).until(EC.visibility_of(input_element)).send_keys(validateonlyname1)
                    time.sleep(0.5)
                elif source_type == "Processes Cat.1":
                    validateonlyname2 = validateOnlyName2[index % len(validateOnlyName2)]
                    WebDriverWait(driver, 10).until(EC.visibility_of(input_element)).send_keys(validateonlyname2)
                    time.sleep(0.5)
                else:
                    validateonlyname3 = validateOnlyName3[index % len(validateOnlyName3)]
                    WebDriverWait(driver, 10).until(EC.visibility_of(input_element)).send_keys(validateonlyname3)
                    time.sleep(0.5)

                    
                    
            elif input_element.get_attribute("placeholder") == "Please enter" and input_element.get_attribute("id") == 'validateOnly_ActivityIntensityUnit': #使用量單位
                WebDriverWait(driver, 10).until(EC.visibility_of(input_element)).send_keys("Kg")
            elif input_element.get_attribute("placeholder") == "Please enter" and input_element.get_attribute("id") == 'validateOnly_Unit': #係數單位
                WebDriverWait(driver, 10).until(EC.visibility_of(input_element)).send_keys("Kg")
                
            elif input_element.get_attribute("placeholder") == "Please enter" and input_element.get_attribute("id") == 'validateOnly_Evidence': #使用量佐證說明
                WebDriverWait(driver, 10).until(EC.visibility_of(input_element)).send_keys("Usage support test")
                
            elif input_element.get_attribute("placeholder") == "Please enter" and input_element.get_attribute("id") == 'validateOnly_Description': #備註
                WebDriverWait(driver, 10).until(EC.visibility_of(input_element)).send_keys("Note test")
                

            elif input_element.get_attribute("id") == 'validateOnly_ingredientName':
                continue

            elif input_element.get_attribute("placeholder") == "Please enter the number" and input_element.get_attribute("id") == 'validateOnly_Scalar': #使用量/使用冷媒/製冷劑填充量
                WebDriverWait(driver, 10).until(EC.visibility_of(input_element)).send_keys(100)
                
            elif input_element.get_attribute("placeholder") == "Please enter the number" and input_element.get_attribute("id") == 'validateOnly_AnnualPurchaseAmount':
                WebDriverWait(driver, 10).until(EC.visibility_of(input_element)).send_keys( ''.join(random.choices(string.digits, k=3)))
                
            elif input_element.get_attribute("placeholder") == "Please enter the number" and input_element.get_attribute("id") == 'validateOnly_OriginalFill': #消防原始填充量
                WebDriverWait(driver, 10).until(EC.visibility_of(input_element)).send_keys(100)
                
            elif input_element.get_attribute("placeholder") == "Please enter the number" and input_element.get_attribute("id") == 'validateOnly_TransportDistance':
                WebDriverWait(driver, 10).until(EC.visibility_of(input_element)).send_keys( ''.join(random.choices(string.digits, k=3)))
                
            elif input_element.get_attribute("placeholder") == "Please enter the number" and input_element.get_attribute("id") == 'validateOnly_TotalNumber':
                WebDriverWait(driver, 10).until(EC.visibility_of(input_element)).send_keys(10)
                
            elif input_element.get_attribute("placeholder") == "Please enter the number" and input_element.get_attribute("id") == 'validateOnly_UsedMonth': #使用月數
                WebDriverWait(driver, 10).until(EC.visibility_of(input_element)).send_keys(12)
                
            elif input_element.get_attribute("placeholder") == "Please enter the number" and input_element.get_attribute("id") == 'validateOnly_Quantity': #消防數量/其他關注類數量
                WebDriverWait(driver, 10).until(EC.visibility_of(input_element)).send_keys(10)
                
           
            elif input_element.get_attribute("placeholder") == "Please enter":
                WebDriverWait(driver, 10).until(EC.visibility_of(input_element)).send_keys("Kg")
                
            elif input_element.get_attribute("type")=="search":   #選單類 先click 再選 再click
        
                if input_element.get_attribute("aria-activedescendant")=="validateOnly_Type_list_0": #車輛類別
                    element = WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Company owned vehicles')]")))
                    element.click()
                elif input_element.get_attribute("id")=="validateOnly_Area": #地區
                    input_element.click()
                    # area_plate = Area[index % len(Area)]
                    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'Taiwan')]")))
                    element.click()
                
                    try:
                        element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'mobile migration source gasoline KgCO2e / L')]"))) #移動源係數
                        element.click()
                    except:
                        try:
                            element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'stationary Stationary gasoline KgCO2e/L')]"))) #固定源係數
                            element.click()
                
                        except:
                            element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'Anthracite KgCO2e/Kg')]"))) #製程係數
                            element.click()
                            time.sleep(0.5)
                    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'2023')]")))
                    element.click()
                    time.sleep(0.5)
                  
                    
                elif input_element.get_attribute("id")=="validateOnly_activityDataType": #活動數據種類
                    input_element.click()
                    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Automatic Continuous Measurement')]")))
                    element.click()
                
                elif input_element.get_attribute("id")=="validateOnly_emitParaType": #排放係數種類
                    input_element.click()
                    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Mass Energy Balance')]")))
                    element.click()
                
                elif input_element.get_attribute("id") == "validateOnly_ARnGWPid": 
                    input_element.click()
                    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Carbon dioxide')]")))
                    element.click()
                
                
                elif input_element.get_attribute("id") == "validateOnly_Commuting":
                    input_element.click()
                    element1 = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH,'//div[@class="ant-select-item-option-content"][1]')))
                    element1.click()     
                    
                elif input_element.get_attribute("id") == "validateOnly_TransportTypeID":
                    input_element.click()
                    element1 = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH,'//div[@class="ant-select-item ant-select-item-option ant-pro-filed-search-select-option "][1]')))
                    element1.click()      
                
                elif input_element.get_attribute("id") == "validateOnly_warmGasType":
                    input_element.click()
                    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '固定燃燒 Stationary combustion')]")))
                    element.click()
                
                elif input_element.get_attribute("id") == "validateOnly_ParameterID":
                    input_element.click()
                    try:
                        validateonlyparameterID = validateOnlyParameterID[index % len(validateOnlyParameterID)]
                        element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(text(), '{validateonlyparameterID}')]"))) #使用冷媒/製冷劑種類
                        element.click()
                
                    except:
                        validateonlyparameterID1 = validateOnlyParameterID1[index % len(validateOnlyParameterID1)]
                        element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(text(), '{validateonlyparameterID1}')]")))
                        element.click()
                
                elif input_element.get_attribute("id") == "validateOnly_ParameterID2": #設備類型(排放因子)
                    input_element.click()
                    validateonlyparameterID2 = validateOnlyParameterID2[index % len(validateOnlyParameterID2)]
                    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(text(), '{validateonlyparameterID2}')]")))
                    element.click()
                
                else:
                    continue  
                
            
        except : #StaleElementReferenceException
            form_element = driver.find_element(By.ID, "validateOnly")               #表單element
            input_elements = form_element.find_elements(By.TAG_NAME, "input")
            if source_type == "Others GHG Cat.1":
                WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "validateOnly_Unit"))).send_keys("t")
                
                WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "validateOnly_ActivityIntensityUnit"))).send_keys("Kg")

    button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']")))    #送出
    button.click()
    time.sleep(2)
    if rfg_or_fire == 1:
        catch_response(driver,'Fugitive Cat.19(RFG)')
    elif rfg_or_fire == 2:
        catch_response(driver,'Fugitive Cat.1(Fire)')
    elif rfg_or_fire == 3:
        catch_response(driver, 'US Oil & electricity Cat.4(Upstream emissions of electricity imported)')
    elif rfg_or_fire == 4:
        catch_response(driver, 'US Oil & electricity Cat.4(Upstream emissions of other energy imported)')
    elif rfg_or_fire == 0:
        catch_response(driver,source_type)
    time.sleep(2)


def workhour(driver,url,index):
    personnel = ["Tester 1","Tester 2","Tester 3"]
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    element.click()
    time.sleep(1)
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//input[@placeholder="Please enter  People type"]')))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//input[@placeholder="Please enter  People type"]')))
    personnel_plate = personnel[index % len(personnel)]
    element.send_keys(personnel_plate)
    time.sleep(1)
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), 'Add')]")))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Add')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(1)

    actions = ActionChains(driver)
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), 'Edit')]")))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Edit')]")))
    actions.move_to_element(element).click().perform()
    time.sleep(1)

    input_elements = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.XPATH, '//input[@placeholder="Please enter"]'))
    )
    # 輸入不同的值到每個 <input> 元素
    for i,element in enumerate(input_elements):
        if i<96:
            #value_to_input = random.randint(1, 20)
            element.send_keys(12)
        else:
            value_to_input = "Test instruction"
            element.send_keys(value_to_input)
    time.sleep(1) 

    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"][@name="ActivityDataType"]')))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"][@name="ActivityDataType"]')))
    element.click()
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Automatic Continuous Measurement')]")))
    element.click()
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"][@name="EmitParaType"]')))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"][@name="EmitParaType"]')))
    element.click()
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Mass Energy Balance')]")))
    element.click()
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), 'Save')]")))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Save')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(3)
    catch_response(driver,'Fugitive Cat.1(Calculate working hours)')

def elec(driver,url,index):
    personnel = ["Tester 1","Tester 2","Tester 3"]
    area = ["Taiwan","中國","緬甸","越南","墨西哥"]
    time.sleep(1)
    page_height = driver.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );")
    # 往下滾動至頁面中間
    driver.execute_script(f"window.scrollTo(0, {page_height // 3});")
    time.sleep(1)
    element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    element.click()
    time.sleep(1)
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//input[@placeholder="Please enter  Electricity No./User No."]')))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//input[@placeholder="Please enter  Electricity No./User No."]')))
    personnel_plate = personnel[index % len(personnel)]
    element.send_keys(personnel_plate)
    time.sleep(1)
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), 'Add')]")))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Add')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(1)
    
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), 'Edit')]")))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Edit')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(1)

    input_elements = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.XPATH, '//input[@placeholder="Please enter"]'))
    )
    # 輸入不同的值到每個 <input> 元素
    for i,element in enumerate(input_elements):
        if i<78:
            value_to_input = random.randint(1, 20)
            element.send_keys(Keys.CONTROL, 'a')   #crtl+A 全選
            element.send_keys(Keys.BACKSPACE)      #刪除
            element.send_keys(100)
        else:
            value_to_input ="Test instruction"
            element.send_keys(value_to_input)
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="ant-select ant-cascader ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-cascader ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    element.click()
    area_plate = area[index % len(area)]
    element = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(text(),'{area_plate}')]")))
    element.click()
    element = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'2023')]")))
    element.click()
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-status-success ant-select-single ant-select-show-arrow"][@name="ActivityDataType"]')))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-status-success ant-select-single ant-select-show-arrow"][@name="ActivityDataType"]')))
    element.click()
    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Automatic Continuous Measurement')]")))
    element.click()
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-status-success ant-select-single ant-select-show-arrow"][@name="EmitParaType"]')))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-status-success ant-select-single ant-select-show-arrow"][@name="EmitParaType"]')))
    element.click()
    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Mass Energy Balance')]")))
    element.click()
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), 'Save')]")))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Save')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(3)  #等他儲存完畢
    catch_response(driver,'Electricity Cat.2(Normal-use electricity)')   #抓response

def green_elec(driver,url):
    element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//div[text()="綠電 B.3.2.a"]')))
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="綠電 B.3.2.a"]')))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(1)
    page_height = driver.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );")
    # 往下滾動至頁面中間
    driver.execute_script(f"window.scrollTo(0, {page_height // 3});")
    time.sleep(1)
    element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '編輯')]")))
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '編輯')]")))
    driver.execute_script("arguments[0].click()", element)
    
    input_elements = WebDriverWait(driver, 5).until(
        EC.presence_of_all_elements_located((By.XPATH, '//input[@placeholder="請輸入"]'))
    )
    
    # 輸入不同的值到每個 <input> 元素
    for i,element in enumerate(input_elements):
        if i<12:
            value_to_input = random.randint(1, 20)
            element.send_keys(Keys.CONTROL, 'a')   #crtl+A 全選
            element.send_keys(Keys.BACKSPACE)      #刪除
            element.send_keys(200)
        else:
            value_to_input = '測試說明'
            element.send_keys(Keys.CONTROL, 'a')   #crtl+A 全選
            element.send_keys(Keys.BACKSPACE)      #刪除
            element.send_keys(value_to_input)

    try:
        element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"][@name="ActivityDataType"]')))
        element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"][@name="ActivityDataType"]')))
        element.click()
        element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '自動連續量測')]")))
        element.click()
        element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"][@name="EmitParaType"]')))
        element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"][@name="EmitParaType"]')))
        element.click()
        element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '自我/量測/質能')]")))
        element.click()
    except:
        element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-status-success ant-select-single ant-select-show-arrow"][@name="ActivityDataType"]')))
        element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-status-success ant-select-single ant-select-show-arrow"][@name="ActivityDataType"]')))
        element.click()
        element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '自動連續量測')]")))
        element.click()
        element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-status-success ant-select-single ant-select-show-arrow"][@name="EmitParaType"]')))
        element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-status-success ant-select-single ant-select-show-arrow"][@name="EmitParaType"]')))
        element.click()
        element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '自我/量測/質能')]")))
        element.click()
        time.sleep(0.5)
    element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '儲存')]")))
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '儲存')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(1)
    catch_response(driver,'輸入電力 C2(綠電)')   #抓response

def steam(driver,url,index):
    time.sleep(1)
    steam_input(driver,url,0,index)   #做蒸氣加項

def steam_input(driver,url,rfg_or_fire,index):
    personnel = ["測試人員1","測試人員2","測試人員3"]
    element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    element.click()
    time.sleep(1)
    element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//input[@placeholder="請輸入  蒸氣編號"]')))
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//input[@placeholder="請輸入  蒸氣編號"]')))
    personnel_plate = personnel[index % len(personnel)]
    element.send_keys(personnel_plate)
    time.sleep(0.5)
    element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '新增')]")))
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '新增')]")))
    driver.execute_script("arguments[0].click()", element)
    
    time.sleep(1)
    page_height = driver.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );")
    # 往下滾動至頁面中間
    driver.execute_script(f"window.scrollTo(0, {page_height // 3});")
    time.sleep(1)
    element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '編輯')]")))
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '編輯')]")))
    driver.execute_script("arguments[0].click()", element)

    element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="ant-select ant-cascader ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-cascader ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    element.click()
    element = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'台灣')]")))
    element.click()
    element = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'2023')]")))
    element.click()
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"][@name="ActivityDataType"]')))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"][@name="ActivityDataType"]')))
    element.click()
    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '自動連續量測')]")))
    element.click()
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"][@name="EmitParaType"]')))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"][@name="EmitParaType"]')))
    element.click()
    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '自我/量測/質能')]")))
    element.click()
    input_elements = WebDriverWait(driver, 5).until(
        EC.presence_of_all_elements_located((By.XPATH, '//input[@placeholder="請輸入"]'))
    )
    # 輸入不同的值到每個 <input> 元素
    for i,element in enumerate(input_elements):
        if i<12:
            value_to_input = random.randint(1, 20)
            element.send_keys(Keys.CONTROL, 'a')   #crtl+A 全選
            element.send_keys(Keys.BACKSPACE)      #刪除
            element.send_keys(100)
        else:
            value_to_input = '測試說明'
            element.send_keys(Keys.CONTROL, 'a')   #crtl+A 全選
            element.send_keys(Keys.BACKSPACE)      #刪除
            element.send_keys(value_to_input)
    time.sleep(1) 
    element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '儲存')]")))
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '儲存')]")))
    driver.execute_script("arguments[0].click()", element)    
    time.sleep(1)
    if rfg_or_fire == 0:
        catch_response(driver,'輸入蒸汽 C2(蒸氣加項)')
    elif rfg_or_fire == 1:
        catch_response(driver,'輸入蒸汽 C2(蒸氣減項))')