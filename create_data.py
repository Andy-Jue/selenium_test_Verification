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

def source_input(driver,url):
    source_type = ['移動源 C1','固定燃燒源 C1','工業製程 C1','人為逸散 C1','其他關注類物質 C1','輸入電力 C2','輸入蒸汽 C2','客戶和訪客運輸 C3','員工差旅 C3','員工通勤 C3','營運產生之廢棄物 C4','資本設備 C4','輸入能源上游排放 C4','用水(水資源管理) B.5.2.a','購買商品 C4','上游租賃資產 C4','顧問諮詢、清潔、維護等 C4','上游運輸及配送 C3','下游運輸及配送 C3','廢棄物運輸 C3','下游租賃資產 C5','產品壽命終止階段 C5','產品使用階段 C5','投資 C5', '其它間接排放 C6']
    url = url+"/calculate/" 

    for i in source_type:
        if i=='移動源 C1' or i == '固定燃燒源 C1'or i == '工業製程 C1' or i== '其他關注類物質 C1':
            link = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.LINK_TEXT, i)))
            link = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))  
            link.click()
            
            # element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '新增')]")))
            # element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '新增')]")))
            # driver.execute_script("arguments[0].click()", element)
            # time.sleep(0.5)
            input_data_C1_C2_verify(driver)
            # input_data_C1_C2(driver, i,0)
        
            
        elif i == '人為逸散 C1':
            link = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.LINK_TEXT, i)))
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))  
            link.click()
            time.sleep(0.5)
            workhour(driver,url)
            
            element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//div[text()="冷媒設備 B.2.2.d"]')))
            element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="冷媒設備 B.2.2.d"]')))
            driver.execute_script("arguments[0].click()", element)

            element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '新增')]")))
            element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '新增')]")))
            driver.execute_script("arguments[0].click()", element)

            time.sleep(0.5)
            input_data_C1_C2(driver, i,1)

            element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//div[text()="消防設備 B.2.2.d"]')))
            element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="消防設備 B.2.2.d"]')))
            driver.execute_script("arguments[0].click()", element)

            element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '新增')]")))
            element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '新增')]")))
            driver.execute_script("arguments[0].click()", element)

            time.sleep(0.5)
            input_data_C1_C2(driver, i,2)
        elif i == '輸入能源上游排放 C4':
            link = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.LINK_TEXT, i)))
            link = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))  
            link.click()
            
            element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//div[text()="輸入電力上游 B.5.2.a"]')))
            element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="輸入電力上游 B.5.2.a"]')))
            driver.execute_script("arguments[0].click()", element)

            element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '新增')]")))
            element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '新增')]")))
            driver.execute_script("arguments[0].click()", element)

            time.sleep(1)
            input_data_C3_C6(driver, i,3) 
            element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//div[text()="其他輸入能源上游 B.5.2.a"]')))
            element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="其他輸入能源上游 B.5.2.a"]')))
            driver.execute_script("arguments[0].click()", element)

            element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '新增')]")))
            element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '新增')]")))
            driver.execute_script("arguments[0].click()", element)

            time.sleep(0.5)
            input_data_C3_C6(driver, i,4)
        elif i == '輸入電力 C2':
            link = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.LINK_TEXT, i)))
            link = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))  
            link.click()
            
            element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//div[text()="一般用電 B.3.2.a"]')))
            element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="一般用電 B.3.2.a"]')))
            driver.execute_script("arguments[0].click()", element)
            
            elec(driver,url)
            time.sleep(0.5)
            
            
            driver.get(url+'electricity') 
            element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//div[text()="綠電 B.3.2.a"]')))
            element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="綠電 B.3.2.a"]')))
            driver.execute_script("arguments[0].click()", element)

            time.sleep(0.5)
            green_elec(driver,url)
        elif i == '輸入蒸汽 C2' :
            link = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.LINK_TEXT, i)))
            link = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))  
            link.click()
            steam(driver,url)
            time.sleep(0.5)
                
                
               
        else:
            link = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.LINK_TEXT, i)))
            link = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))  
            link.click()
            
            element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '新增')]")))
            element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '新增')]")))
            driver.execute_script("arguments[0].click()", element)
            time.sleep(0.5)
            input_data_C3_C6(driver, i,0)
        
        driver.get(url)
    # driver.quit()
def coefficient(driver):
    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "validateOnly")))
    button = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//button[@class='ant-btn css-dts6b9 ant-btn-link d-flex align-items-center mt-2']")))
    button.click()
    time.sleep(1)
    element = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//div[@class='ant-collapse-header'][1]")))
    element.click()
    time.sleep(1)
    random_index = random.randint(1, 17)
    element = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//div[@class='ant-collapse-item'][1]")))
    element.click()
    time.sleep(1)
    element = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//a[@class='d-block my-2 text-dark'][1]")))
    element.click()
    time.sleep(1)
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//tr[@class='ant-table-row ant-table-row-level-0'][1]//span[@class='ant-checkbox-inner']")))
    actions = ActionChains(driver)
    actions.move_to_element(element).click().perform()
    time.sleep(1)
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//button[@class='ant-btn css-dts6b9 ant-btn-primary']//span[contains(text(),'確 定')]")))
    element.click()
    time.sleep(1)


def input_data_C3_C6(driver, source_type,rfg_or_fire):
    # random_letters = ''.join(random.choices(string.ascii_letters, k=10))
    
    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "validateOnly")))
    time.sleep(2)
    coefficient(driver)
    time.sleep(1)
    

    form_element = driver.find_element(By.ID, "validateOnly")               #表單element

    input_elements = form_element.find_elements(By.TAG_NAME, "input")       #tag_name是input的
    
    new_inputelements =[]

    for i in input_elements:                     #把disable的先移除掉
        if i.get_attribute("disabled") != None and i.get_attribute("id")!="validateOnly_referenceFile" and i.get_attribute("type")!="search":
            print("")
        else:
            new_inputelements.append(i)

    time.sleep(2)
    for input_element in new_inputelements:
        try:
            if input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute("id") == 'validateOnly_GHGEvaluateItem': #評估項目
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("評估項目測試")
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute("id") == 'validateOnly_MaterialNo': #料號
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("料號測試")
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute("id") == 'validateOnly_ActivityIntensityUnit': #活動強度(使用量)單位
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("Kg")
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute("id") == 'validateOnly_MaterialName': #物料名稱
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("物料名稱測試")
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute("id") == 'validateOnly_CargoName': #上游評估項目
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("上游評估項目測試")
                time.sleep(1)
            
            elif input_element.get_attribute("id") == 'validateOnly_ingredientName':
                continue

            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_ActivityIntensity': #活動強度(使用量)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(100)
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_AnnualPurchaseAmount': #年度採購量
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(100)
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_TransportWeight': #上游重量
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(10)
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_TransportDistance': #上游運輸距離(km)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(10)
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_TotalEmployees':#同行人數
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(10)
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_WorkingDays': #盤查期間次數
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(24)
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_AverageMovingDistance': #移動距離
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(100)
                time.sleep(1)
   
            elif input_element.get_attribute("type")=="search":   #選單類 先click 再選 再click
        
                if input_element.get_attribute("id")=="validateOnly_activityDataType":            
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '自動連續')]")))
                    element.click()
                    time.sleep(1)
                elif input_element.get_attribute("id")=="validateOnly_emitParaType":
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '自我/量測')]")))
                    element.click()
                    time.sleep(1)
                elif input_element.get_attribute("id") == "validateOnly_DataSource":
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '全球')]")))
                    element.click()
                    time.sleep(1)
                elif input_element.get_attribute("id") == "validateOnly_DataAttribute":
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '基於')]")))
                    element.click()
                    time.sleep(1)
                elif input_element.get_attribute("id") == "validateOnly_Commuting": #員工差旅交通方式
                    input_element.click()
                    element1 = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH,'//div[@class="ant-select-item-option-content"][1]')))
                    element1.click()     
                    time.sleep(1)
                elif input_element.get_attribute("id") == "validateOnly_TransportTypeID": #運輸方式
                    input_element.click()
                    element1 = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH,'//div[@class="ant-select-item ant-select-item-option ant-pro-filed-search-select-option "][1]')))
                    element1.click()      
                    time.sleep(1)
                else:
                    continue  
                time.sleep(1)
            
        except : #StaleElementReferenceException
            print("")
            
                    
    button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']")))    #送出
    button.click()
    time.sleep(3)
    if rfg_or_fire == 1:
        catch_response(driver,'人為逸散 C1(冷媒設備)')
    elif rfg_or_fire == 2:
        catch_response(driver,'人為逸散 C1(消防設備)')
    elif rfg_or_fire == 3:
        catch_response(driver, '輸入能源上游排放 C4(輸入電力上游)')
    elif rfg_or_fire == 4:
        catch_response(driver, '輸入能源上游排放 C4(其他輸入能源上游)')
    elif rfg_or_fire == 0:
        catch_response(driver,source_type)
    time.sleep(2)

def input_data_C1_C2_verify(driver):
    try:
        table = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, "//div[@class='ant-table-content']")))
        print(table.text)
    except:
        print("沒有")
    
    
    


def input_data_C1_C2(driver, source_type,rfg_or_fire):
    # random_letters = ''.join(random.choices(string.ascii_letters, k=10))
    
    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "validateOnly")))
    
    form_element = driver.find_element(By.ID, "validateOnly")               #表單element

    input_elements = form_element.find_elements(By.TAG_NAME, "input")       #tag_name是input的
    
    new_inputelements =[]

    for i in input_elements:                     #把disable的先移除掉
        if i.get_attribute("disabled") != None and i.get_attribute("id")!="validateOnly_referenceFile" and i.get_attribute("type")!="search":
            print("")
        else:
            new_inputelements.append(i)
    for input_element in new_inputelements:
        try:
            if input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute("id") == 'validateOnly_CarPlateNo': #車牌
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("ABC-123")
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute("id") == 'validateOnly_ProcessName': #製程別
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("製程別測試")
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute("id") == 'validateOnly_ResponsibleUnit': #負責單位
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("負責單位測試")
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute("id") == 'validateOnly_Name': #公務車型/設備名稱
                if source_type == "移動源 C1":
                    WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("公務車型測試")
                    time.sleep(0.5)
                elif source_type == "固定燃燒源 C1":
                    WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("固定設備名稱測試")
                    time.sleep(0.5)
                elif source_type == "工業製程 C1":
                    WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("製程設備名稱測試")
                    time.sleep(0.5)
                else:
                    WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("設備名稱測試")
                    time.sleep(0.5)
                    
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute("id") == 'validateOnly_ActivityIntensityUnit': #使用量單位
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("Kg")
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute("id") == 'validateOnly_Evidence': #使用量佐證說明
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("使用量佐證說明測試")
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute("id") == 'validateOnly_Description': #備註
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("備註測試")
                time.sleep(1)

            elif input_element.get_attribute("id") == 'validateOnly_ingredientName':
                continue

            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_Scalar': #使用量/使用冷媒/製冷劑填充量
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(100)
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_AnnualPurchaseAmount':
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys( ''.join(random.choices(string.digits, k=3)))
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_TransportWeight':
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys( ''.join(random.choices(string.digits, k=3)))
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_TransportDistance':
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys( ''.join(random.choices(string.digits, k=3)))
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_TotalNumber':
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(10)
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_UsedMonth': #使用月數
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(12)
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_Quantity': #消防數量/其他關注類數量
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(10)
                time.sleep(1)
           
            elif input_element.get_attribute("placeholder") == "請輸入活動強度單位":
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("Kg")
                
            elif input_element.get_attribute("type")=="search":   #選單類 先click 再選 再click
        
                if input_element.get_attribute("aria-activedescendant")=="validateOnly_Type_list_0": #車輛類別
                    element = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '公務車')]")))
                    element.click()
                elif input_element.get_attribute("id")=="validateOnly_Area": #地區
                    input_element.click()
                    element = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'台灣')]")))
                    element.click()
                    time.sleep(1)
                    try:
                        element = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'移動源 - 汽油')]"))) #移動源係數
                        element.click()
                    except:
                        try:
                            element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'固定源 - 汽油 ')]"))) #固定源係數
                            element.click()
                            time.sleep(0.5)
                        except:
                            element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'無煙煤 KgCO2e/Kg')]"))) #製程係數
                            element.click()
                            time.sleep(0.5)
                    element = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'2023')]")))
                    element.click()
                    time.sleep(0.5)
                  
                    
                elif input_element.get_attribute("id")=="validateOnly_activityDataType": #活動數據種類
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '自動連續')]")))
                    element.click()
                    time.sleep(1)
                elif input_element.get_attribute("id")=="validateOnly_emitParaType": #排放係數種類
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '自我/量測')]")))
                    element.click()
                    time.sleep(1)
                elif input_element.get_attribute("id") == "validateOnly_ARnGWPid": 
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Carbon dioxide')]")))
                    element.click()
                    time.sleep(1)  
                elif input_element.get_attribute("id") == "validateOnly_DataSource":
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '全球')]")))
                    element.click()
                    time.sleep(1)
                elif input_element.get_attribute("id") == "validateOnly_DataAttribute":
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '基於')]")))
                    element.click()
                    time.sleep(1)
                elif input_element.get_attribute("id") == "validateOnly_Commuting":
                    input_element.click()
                    element1 = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH,'//div[@class="ant-select-item-option-content"][1]')))
                    element1.click()     
                    time.sleep(1)
                elif input_element.get_attribute("id") == "validateOnly_RoundTrip":
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '往返')]")))
                    element.click()  
                    time.sleep(1)
                elif input_element.get_attribute("id") == "validateOnly_TransportTypeID":
                    input_element.click()
                    element1 = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH,'//div[@class="ant-select-item ant-select-item-option ant-pro-filed-search-select-option "][1]')))
                    element1.click()      
                    time.sleep(1)
                elif input_element.get_attribute("id") == "validateOnly_warmGasType":
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '固定燃燒 Stationary combustion')]")))
                    element.click()
                    time.sleep(1)
                elif input_element.get_attribute("id") == "validateOnly_DeviceName_SystemMenusID":
                    if source_type=='移動源 C1':
                        input_element.click()
                        element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//div[@label="公務車(燃料單據)"]')))
                        element.click()
                        time.sleep(1)
                    if source_type=='固定燃燒源 C1':
                        input_element.click()
                        element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH,'//div[@label="緊急發電機(燃料單)"]')))
                        driver.execute_script("arguments[0].click()", element)
                        time.sleep(1)
                    if source_type=='工業製程 C1':
                        input_element.click()
                        element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//div[@label="金屬切割器(乙炔用量)"]')))
                        driver.execute_script("arguments[0].click()", element)
                        time.sleep(1)
                elif input_element.get_attribute("id") == "validateOnly_ParameterID":
                    input_element.click()
                    try:
                        element = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'R-11')]"))) #使用冷媒/製冷劑種類
                        element.click()
                        time.sleep(1)
                    except:
                        element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'R-123')]")))
                        element.click()
                        time.sleep(1)
                elif input_element.get_attribute("id") == "validateOnly_ParameterID2": #設備類型(排放因子)
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '家用冰箱')]")))
                    element.click()
                    time.sleep(1)
                else:
                    continue  
                time.sleep(1)
            
        except : #StaleElementReferenceException
            form_element = driver.find_element(By.ID, "validateOnly")               #表單element
            input_elements = form_element.find_elements(By.TAG_NAME, "input")
            if source_type == "其他關注類物質 C1":
                WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "validateOnly_Unit"))).send_keys("t")
                time.sleep(1)
                WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "validateOnly_ActivityIntensityUnit"))).send_keys("Kg")
                time.sleep(1)

    button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']")))    #送出
    button.click()
    time.sleep(3)
    if rfg_or_fire == 1:
        catch_response(driver,'人為逸散 C1(冷媒設備)')
    elif rfg_or_fire == 2:
        catch_response(driver,'人為逸散 C1(消防設備)')
    elif rfg_or_fire == 3:
        catch_response(driver, '輸入能源上游排放 C4(輸入電力上游)')
    elif rfg_or_fire == 4:
        catch_response(driver, '輸入能源上游排放 C4(其他輸入能源上游)')
    elif rfg_or_fire == 0:
        catch_response(driver,source_type)
    time.sleep(2)

def workhour(driver,url):
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    element.click()
    time.sleep(2)
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//input[@placeholder="請輸入  人員類別"]')))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//input[@placeholder="請輸入  人員類別"]')))
    element.send_keys("測試人員1")
    time.sleep(2)
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '新增')]")))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '新增')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(3)

    actions = ActionChains(driver)
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '編輯')]")))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '編輯')]")))
    actions.move_to_element(element).click().perform()
    time.sleep(2)

    input_elements = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.XPATH, '//input[@placeholder="請輸入"]'))
    )
    # 輸入不同的值到每個 <input> 元素
    for i,element in enumerate(input_elements):
        if i<96:
            #value_to_input = random.randint(1, 20)
            element.send_keys(12)
        else:
            value_to_input = "測試說明"
            element.send_keys(value_to_input)
    time.sleep(3) 

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
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '儲存')]")))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '儲存')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(5)
    catch_response(driver,'人為逸散 C1(工時計算)')

def elec(driver,url):
    driver.get(url+'electricity')  #driver到電力頁面
    time.sleep(0.5)
    page_height = driver.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );")
    # 往下滾動至頁面中間
    driver.execute_script(f"window.scrollTo(0, {page_height // 3});")
    time.sleep(0.5)
    element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    element.click()
    time.sleep(1)
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//input[@placeholder="請輸入  電號/用戶編號"]')))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//input[@placeholder="請輸入  電號/用戶編號"]')))
    element.send_keys("測試人員1")
    time.sleep(1)
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '新增')]")))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '新增')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(0.5)
    
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '編輯')]")))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '編輯')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(1)
    
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="ant-select ant-cascader ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-cascader ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    element.click()
    element = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'台灣')]")))
    element.click()
    element = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'2023')]")))
    element.click()
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
    input_elements = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.XPATH, '//input[@placeholder="請輸入"]'))
    )
    # 輸入不同的值到每個 <input> 元素
    for i,element in enumerate(input_elements):
        if i<78:
            value_to_input = random.randint(1, 20)
            element.send_keys(Keys.CONTROL, 'a')   #crtl+A 全選
            element.send_keys(Keys.BACKSPACE)      #刪除
            element.send_keys(100)
        else:
            value_to_input ="測試說明"
            element.send_keys(value_to_input)
    time.sleep(2) 
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '儲存')]")))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '儲存')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(3)  #等他儲存完畢
    catch_response(driver,'輸入電力 C2(一般用電)')   #抓response

def green_elec(driver,url):
    page_height = driver.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );")
    # 往下滾動至頁面中間
    driver.execute_script(f"window.scrollTo(0, {page_height // 3});")
    time.sleep(1)
    element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//div[text()="綠電 B.3.2.a"]')))
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="綠電 B.3.2.a"]')))
    driver.execute_script("arguments[0].click()", element)
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

    time.sleep(1) 
    element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '儲存')]")))
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '儲存')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(1)
    catch_response(driver,'輸入電力 C2(綠電)')   #抓response

def steam(driver,url):
    time.sleep(1)
    steam_input(driver,url,0)   #做蒸氣加項

def steam_input(driver,url,rfg_or_fire):

    element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    element.click()

    element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//input[@placeholder="請輸入  蒸氣編號"]')))
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//input[@placeholder="請輸入  蒸氣編號"]')))
    element.send_keys("測試人員1")
    time.sleep(0.5)
    element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '新增')]")))
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '新增')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(1)
    page_height = driver.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );")
    # 往下滾動至頁面中間
    driver.execute_script(f"window.scrollTo(0, {page_height // 3});")

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