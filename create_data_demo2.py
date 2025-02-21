import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import random
from response import catch_response
import string

def source_input_demo2(driver,url):
    source_type = ['移動源 C1', '固定燃燒源 C1', '工業製程 C1', '人為逸散 C1', '其他關注類物質 C1', '輸入電力 C2',
                   '客戶和訪客運輸 C3', '員工差旅 C3', '員工通勤 C3', '營運產生之廢棄物 C4', '資本設備 C4',
                   '輸入能源上游排放 C4', '用水(水資源管理) B.5.2.a', '購買商品 C4', '上游租賃資產 C4',
                   '顧問諮詢、清潔、維護等 C4', '上游運輸及配送 C3', '廢棄物運輸 C3', '下游運輸及配送 C3',
                   '下游租賃資產 C5', '產品壽命終止階段 C5', '產品使用階段 C5', '投資 C5', '其它間接排放 C6']
    url = url+"/project/calculate/"

    for i in source_type:
        if i=='移動源 C1' or i == '固定燃燒源 C1'or i == '工業製程 C1' or i== '其他關注類物質 C1':
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))
            link.click()
            for j in range(3):
                element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '新增')]")))
                driver.execute_script("arguments[0].click()", element)
                time.sleep(1)
                input_data_C1_C2(driver, i,j,0)

                
        
            
        elif i == '人為逸散 C1':
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))
            link.click()
            time.sleep(1)
            for j in range(1):
                workhour(driver,url,j)
                time.sleep(1)
                
            element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="冷媒設備 B.2.2.d"]')))
            driver.execute_script("arguments[0].click()", element)
            for j in range(3):
                element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '新增')]")))
                driver.execute_script("arguments[0].click()", element)
                time.sleep(1)
                input_data_C1_C2(driver, i,j,1)
                time.sleep(1)

            element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="消防設備 B.2.2.d"]')))
            driver.execute_script("arguments[0].click()", element)
            for j in range(3):
                element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '新增')]")))
                driver.execute_script("arguments[0].click()", element)
                time.sleep(1)
                input_data_C1_C2(driver, i,j,2)
                time.sleep(1)
        elif i == '輸入能源上游排放 C4':
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))
            link.click()

            for j  in range(3):              
                element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '新增')]")))
                driver.execute_script("arguments[0].click()", element)    
                time.sleep(1)
                input_data_C3_C6(driver, i,j,3) 
                time.sleep(1)
            
            element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="其他輸入能源上游 B.5.2.a"]')))
            driver.execute_script("arguments[0].click()", element)
            for j in range(3):
                element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '新增')]")))
                driver.execute_script("arguments[0].click()", element)
                time.sleep(1)
                input_data_C3_C6(driver, i,j,4)
        elif i == '輸入電力 C2':
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))
            link.click()
            for j in range(1):
                elec(driver,url,j)
                time.sleep(1)
            #     driver.get(url+"electricity")
            #     time.sleep(0.5)
                
            # time.sleep(0.5)
            # green_elec(driver,url)
        elif i == '輸入蒸汽 C2' :
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))
            link.click()
            for j in range(1):
                time.sleep(1)
                steam(driver,url,j)
                time.sleep(1)
                
        else:
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))
            link.click()
            
            for j in range(3):
                element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '新增')]")))
                driver.execute_script("arguments[0].click()", element)
                time.sleep(1)
                input_data_C3_C6(driver, i,j,0)
                time.sleep(1)
        
        driver.get(url)
def wait_and_click(driver, by, value, wait_time=2):
    element = WebDriverWait(driver, wait_time).until(EC.element_to_be_clickable((by, value)))
    element.click()
    time.sleep(1)  # 若需要等待一段時間
    
def coefficient_C4_C6(driver):
    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "validateOnly")))
    time.sleep(2)
    wait_and_click(driver, By.XPATH, "//button[@class='ant-btn css-dts6b9 ant-btn-link d-flex align-items-center mt-2']")
    
    wait_and_click(driver, By.XPATH, "//div[@class='ant-collapse-item'][1]")
    
    wait_and_click(driver, By.XPATH, "//div[@class='ant-collapse-item']//span[contains(text(), '1# ELECTRONIC (918)')]")
    
    wait_and_click(driver, By.XPATH, "//a[@class='d-block my-2 text-dark'][1]")

    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//tr[@class='ant-table-row ant-table-row-level-0'][1]//span[@class='ant-checkbox-inner']")))
    actions = ActionChains(driver)
    actions.move_to_element(element).click().perform()
    time.sleep(1)
    
    wait_and_click(driver, By.XPATH, "//button[@class='ant-btn css-dts6b9 ant-btn-primary']//span[contains(text(),'確 定')]")

def coefficient_C3(driver):
    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "validateOnly")))
    time.sleep(2)
    wait_and_click(driver, By.XPATH, "//button[@class='ant-btn css-dts6b9 ant-btn-link d-flex align-items-center mt-2']")
    
    wait_and_click(driver, By.XPATH, "//div[@class='ant-collapse-item'][1]")
    
    wait_and_click(driver, By.XPATH, "//div[@class='ant-collapse-item']//span[contains(text(), '13# TRANSPORTATION(1993)')]")
    
    wait_and_click(driver, By.XPATH, "//a[@class='d-block my-2 text-dark'][1]")

    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//tr[@class='ant-table-row ant-table-row-level-0'][1]//span[@class='ant-checkbox-inner']")))
    actions = ActionChains(driver)
    actions.move_to_element(element).click().perform()
    time.sleep(1)
    
    wait_and_click(driver, By.XPATH, "//button[@class='ant-btn css-dts6b9 ant-btn-primary']//span[contains(text(),'確 定')]")
def coefficient_C4(driver):
    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "validateOnly")))
    time.sleep(2)
    wait_and_click(driver, By.XPATH, "//button[@class='ant-btn css-dts6b9 ant-btn-link d-flex align-items-center mt-2']")
    
    wait_and_click(driver, By.XPATH, "//div[@class='ant-collapse-item'][1]")
    
    wait_and_click(driver, By.XPATH, "//div[@class='ant-collapse-item']//span[contains(text(), '15#WASTE TREATMENT(5086)')]")
    
    wait_and_click(driver, By.XPATH, "//a[@class='d-block my-2 text-dark'][1]")

    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//tr[@class='ant-table-row ant-table-row-level-0'][1]//span[@class='ant-checkbox-inner']")))
    actions = ActionChains(driver)
    actions.move_to_element(element).click().perform()
    time.sleep(1)
    
    wait_and_click(driver, By.XPATH, "//button[@class='ant-btn css-dts6b9 ant-btn-primary']//span[contains(text(),'確 定')]")


def input_data_C3_C6(driver, source_type,index,rfg_or_fire):
    validateOnlyCommuting = ["其他","汽車-汽油","汽車-柴油"]
    validateOnlyGHGEvaluateItem = ["評估項目測試1","評估項目測試2","評估項目測試3","評估項目測試4","評估項目測試5"]
    validateOnlyMaterialNo = ["料號測試1","料號測試2","料號測試3","料號測試4","料號測試5"]
    validateOnlyMaterialName = ["物料名稱測試1","物料名稱測試2","物料名稱測試3","物料名稱測試4","物料名稱測試5"]
    validateOnlyCargoName = ["上游評估項目測試1","上游評估項目測試2","上游評估項目測試3","上游評估項目測試4","上游評估項目測試5"]
    validateOnlyTransportTypeID = ["陸運,運輸 貨車 3.5-7.5噸 EURO5/RER S ","陸運,貨車 3.5-7.5噸","陸運,輕型商用車","陸運,無區分","陸運,貨車 7.5-16噸-冷藏"]
    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "validateOnly")))
    source_types_to_check = {
                            "客戶和訪客運輸 C3", 
                            "員工差旅 C3", 
                            "員工通勤 C3", 
                            "上游運輸及配送 C3", 
                            "廢棄物運輸 C3", 
                            "下游運輸及配送 C3"
                            }
    if source_type in source_types_to_check:
        coefficient_C3(driver)
        time.sleep(1)
    elif source_type == '營運產生之廢棄物 C4':
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
            if input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute("id") == 'validateOnly_GHGEvaluateItem': #評估項目
                if rfg_or_fire == 3 :
                     validateOnly_GHGEvaluateItem = "輸入電力上游" + validateOnlyGHGEvaluateItem[index % len(validateOnlyGHGEvaluateItem)]
                     WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(validateOnly_GHGEvaluateItem)
                elif rfg_or_fire == 4:
                    validateOnly_GHGEvaluateItem = "輸入能源上游" + validateOnlyGHGEvaluateItem[index % len(validateOnlyGHGEvaluateItem)]
                    WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(validateOnly_GHGEvaluateItem)
                else:
                    validateOnly_GHGEvaluateItem = source_type + validateOnlyGHGEvaluateItem[index % len(validateOnlyGHGEvaluateItem)]
                    WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(validateOnly_GHGEvaluateItem)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute("id") == 'validateOnly_MaterialNo': #料號
                validateOnly_MaterialNo = validateOnlyMaterialNo[index % len(validateOnlyMaterialNo)]
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(validateOnly_MaterialNo)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute("id") == 'validateOnly_ActivityIntensityUnit': #活動強度(使用量)單位
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("Kg")

            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute("id") == 'validateOnly_MaterialName': #物料名稱
                validateOnly_MaterialName = source_type +  validateOnlyMaterialName[index % len(validateOnlyMaterialName)]
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(validateOnly_MaterialName)
                
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute("id") == 'validateOnly_CargoName':  #上游評估項目
                validateOnly_CargoName = validateOnlyCargoName[index % len(validateOnlyCargoName)]
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(validateOnly_CargoName)
                
            
            elif input_element.get_attribute("id") == 'validateOnly_ingredientName':
                continue

            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_ActivityIntensity': #活動強度(使用量)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(100)
                
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_AnnualPurchaseAmount': #年度採購量
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(100)
                
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_TransportWeight': #上游重量
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(10)
                
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_TransportDistance': #上游運輸距離(km)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(10)
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute(
                    "id") == 'validateOnly_TransportQuantity':  # 上游運輸次數
                WebDriverWait(driver, 10).until(EC.visibility_of(input_element)).send_keys(10)
                time.sleep(1)
                
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_TotalEmployees':#同行人數
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(10)
                
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_WorkingDays': #盤查期間次數
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(24)
                
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_AverageMovingDistance': #移動距離
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(100)
                
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_MaterialSpec': #單一物料重量
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(100)
                
                
            elif input_element.get_attribute("placeholder") == "請輸入活動強度單位" and input_element.get_attribute("id") == 'validateOnly_ActivityIntensityUnit': #單一物料重量單位
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("Kg")
                
   
            elif input_element.get_attribute("type")=="search":   #選單類 先click 再選 再click
        
                if input_element.get_attribute("id")=="validateOnly_activityDataType":            
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '自動連續')]")))
                    element.click()
                
                elif input_element.get_attribute("id")=="validateOnly_emitParaType":
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '自我/量測')]")))
                    element.click()
                
                elif input_element.get_attribute("id") == "validateOnly_DataSource":
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '全球')]")))
                    element.click()
                
                elif input_element.get_attribute("id") == "validateOnly_DataAttribute":
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '基於')]")))
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
        catch_response(driver,'人為逸散 C1(冷媒設備)')
    elif rfg_or_fire == 2:
        catch_response(driver,'人為逸散 C1(消防設備)')
    elif rfg_or_fire == 3:
        catch_response(driver, '輸入能源上游排放 C4(輸入電力上游)')
    elif rfg_or_fire == 4:
        catch_response(driver, '輸入能源上游排放 C4(其他輸入能源上游)')
    elif rfg_or_fire == 0:
        catch_response(driver,source_type)
    
    
    


def input_data_C1_C2(driver, source_type,index,rfg_or_fire):

    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "validateOnly")))
    
    form_element = driver.find_element(By.ID, "validateOnly")               #表單element

    input_elements = form_element.find_elements(By.TAG_NAME, "input")       #tag_name是input的
    
    CarPlateNo = ["ABC-123","ABC-234","ABC-345","ABC-456","ABC-567"]
    validateOnlyName1 = ["固定設備名稱測試1","固定設備名稱測試2","固定設備名稱測試3","固定設備名稱測試4","固定設備名稱測試5"]
    validateOnlyName2 = ["製程設備名稱測試1","製程設備名稱測試2","製程設備名稱測試3","製程設備名稱測試4","製程設備名稱測試5"]
    validateOnlyName3 = ["設備名稱測試1","設備名稱測試2","設備名稱測試3","設備名稱測試4","設備名稱測試5"]
    Area = ["台灣","大陸","越南","泰國","墨西哥"]
    validateOnlyParameterID = ["R-11","R-12","R-13","R-112","R-112a"]
    validateOnlyParameterID1 = ["R-123","R-125","FM-200","HFC-236fa","Halon-1201"]
    validateOnlyParameterID2 = ["家用冰箱","飲水機","除濕機","餐廳冷藏櫃","恆溫恆濕機"]
    
    
    new_inputelements =[]

    for i in input_elements:                     #把disable的先移除掉
        if i.get_attribute("disabled") != None and i.get_attribute("id")!="validateOnly_referenceFile" and i.get_attribute("type")!="search":
            print("")
        else:
            new_inputelements.append(i)
    for input_element in new_inputelements:
        try:
            if input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute("id") == 'validateOnly_CarPlateNo': #車牌
                car_plate = CarPlateNo[index % len(CarPlateNo)]
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(str(car_plate))
                
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute("id") == 'validateOnly_ProcessName': #製程別
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("製程別測試")
                
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute("id") == 'validateOnly_ResponsibleUnit': #負責單位
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("負責單位測試")
                
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute("id") == 'validateOnly_Name': #公務車型/設備名稱
                if source_type == "移動源 C1":
                    WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("公務車型測試")
                    time.sleep(0.5)
                elif source_type == "固定燃燒源 C1":
                    validateonlyname1 = validateOnlyName1[index % len(validateOnlyName1)]
                    WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(validateonlyname1)
                    time.sleep(0.5)
                elif source_type == "工業製程 C1":
                    validateonlyname2 = validateOnlyName2[index % len(validateOnlyName2)]
                    WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(validateonlyname2)
                    time.sleep(0.5)
                else:
                    validateonlyname3 = validateOnlyName3[index % len(validateOnlyName3)]
                    WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(validateonlyname3)
                    time.sleep(0.5)
                    
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute("id") == 'validateOnly_ActivityIntensityUnit': #使用量單位
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("Kg")
                
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute("id") == 'validateOnly_Evidence': #使用量佐證說明
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("使用量佐證說明測試")
                
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute("id") == 'validateOnly_Description': #備註
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("備註測試")
                

            elif input_element.get_attribute("id") == 'validateOnly_ingredientName':
                continue

            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_Scalar': #使用量/使用冷媒/製冷劑填充量
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(100)
                
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_AnnualPurchaseAmount':
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys( ''.join(random.choices(string.digits, k=3)))
                
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_OriginalFill': #消防原始填充量
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(100)
                
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_TransportDistance':
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys( ''.join(random.choices(string.digits, k=3)))
                
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_TotalNumber':
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(10)
                
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_UsedMonth': #使用月數
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(12)
                
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_Quantity': #消防數量/其他關注類數量
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(10)
                
           
            elif input_element.get_attribute("placeholder") == "請輸入活動強度單位":
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("Kg")
                
            elif input_element.get_attribute("type")=="search":   #選單類 先click 再選 再click
        
                if input_element.get_attribute("aria-activedescendant")=="validateOnly_Type_list_0": #車輛類別
                    element = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '公務車')]")))
                    element.click()
                elif input_element.get_attribute("id")=="validateOnly_Area": #地區
                    input_element.click()
                    element = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(text(),'台灣')]")))
                    element.click()
                
                    try:
                        element = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'移動源 - 汽油')]"))) #移動源係數
                        element.click()
                    except:
                        try:
                            element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'固定源 - 汽油 ')]"))) #固定源係數
                            element.click()
                
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
                
                elif input_element.get_attribute("id")=="validateOnly_emitParaType": #排放係數種類
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '自我/量測')]")))
                    element.click()
                
                elif input_element.get_attribute("id") == "validateOnly_ARnGWPid": 
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Carbon dioxide')]")))
                    element.click()
                
                elif input_element.get_attribute("id") == "validateOnly_DataSource":
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '全球')]")))
                    element.click()
                
                elif input_element.get_attribute("id") == "validateOnly_DataAttribute":
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '基於')]")))
                    element.click()
                
                elif input_element.get_attribute("id") == "validateOnly_Commuting":
                    input_element.click()
                    element1 = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH,'//div[@class="ant-select-item-option-content"][1]')))
                    element1.click()     
                
                elif input_element.get_attribute("id") == "validateOnly_RoundTrip":
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '往返')]")))
                    element.click()  
                
                elif input_element.get_attribute("id") == "validateOnly_TransportTypeID":
                    input_element.click()
                    element1 = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH,'//div[@class="ant-select-item ant-select-item-option ant-pro-filed-search-select-option "][1]')))
                    element1.click()      
                
                elif input_element.get_attribute("id") == "validateOnly_warmGasType":
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '固定燃燒')]")))
                    element.click()
                
                elif input_element.get_attribute("id") == "validateOnly_ParameterID":
                    input_element.click()
                    try:
                        validateonlyparameterID = validateOnlyParameterID[index % len(validateOnlyParameterID)]
                        element = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(text(), '{validateonlyparameterID}')]"))) #使用冷媒/製冷劑種類
                        element.click()
                
                    except:
                        validateonlyparameterID1 = validateOnlyParameterID1[index % len(validateOnlyParameterID1)]
                        element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(text(), '{validateonlyparameterID1}')]")))
                        element.click()
                
                elif input_element.get_attribute("id") == "validateOnly_ParameterID2": #設備類型(排放因子)
                    input_element.click()
                    validateonlyparameterID2 = validateOnlyParameterID2[index % len(validateOnlyParameterID2)]
                    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(text(), '{validateonlyparameterID2}')]")))
                    element.click()
                
                else:
                    continue  
                
            
        except : #StaleElementReferenceException
            form_element = driver.find_element(By.ID, "validateOnly")               #表單element
            if source_type == "其他關注類物質 C1":
                WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "validateOnly_Unit"))).send_keys("t")
                
                WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "validateOnly_ActivityIntensityUnit"))).send_keys("Kg")


    button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']")))    #送出
    button.click()
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


def workhour(driver,url,index):
    personnel = ["測試人員1","測試人員2","測試人員3"]
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    element.click()
    time.sleep(1)
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//input[@placeholder="請輸入  人員類別"]')))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//input[@placeholder="請輸入  人員類別"]')))
    personnel_plate = personnel[index % len(personnel)]
    element.send_keys(personnel_plate)
    time.sleep(1)
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '新增')]")))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '新增')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(1)

    actions = ActionChains(driver)
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '編輯')]")))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '編輯')]")))
    actions.move_to_element(element).click().perform()
    time.sleep(1)

    input_elements = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.XPATH, '//input[@placeholder="請輸入"]'))
    )
    # 輸入不同的值到每個 <input> 元素
    for i,element in enumerate(input_elements):
        if i<96:
            element.send_keys(12)
        else:
            value_to_input = "測試說明"
            element.send_keys(value_to_input)
    time.sleep(1) 

    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"][@name="ActivityDataType"]')))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"][@name="ActivityDataType"]')))
    element.click()
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '自動連續量測')]")))
    element.click()
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"][@name="EmitParaType"]')))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"][@name="EmitParaType"]')))
    element.click()
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '自我/量測/質能')]")))
    element.click()
    element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '儲存')]")))
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '儲存')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(3)
    catch_response(driver,'人為逸散 C1(工時計算)')

def elec(driver,url,index):
    personnel = ["測試人員1","測試人員2","測試人員3"]
    area = ["台灣","中國","緬甸","越南","墨西哥"]
    time.sleep(1)
    page_height = driver.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );")
    # 往下滾動至頁面中間
    driver.execute_script(f"window.scrollTo(0, {page_height // 3});")
    time.sleep(1)
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    element.click()
    time.sleep(1)
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//input[@placeholder="請輸入  電號/用戶編號"]')))
    personnel_plate = personnel[index % len(personnel)]
    element.send_keys(personnel_plate)
    time.sleep(1)
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '新增')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(1)
    
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '編輯')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(1)

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
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-cascader ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    element.click()
    area_plate = area[index % len(area)]
    element = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(text(),'{area_plate}')]")))
    element.click()
    element = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'2023')]")))
    element.click()
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-status-success ant-select-single ant-select-show-arrow"][@name="ActivityDataType"]')))
    element.click()
    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '自動連續量測')]")))
    element.click()
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-status-success ant-select-single ant-select-show-arrow"][@name="EmitParaType"]')))
    element.click()
    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '自我/量測/質能')]")))
    element.click()
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '儲存')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(3)  #等他儲存完畢
    catch_response(driver,'輸入電力 C2(一般用電)')   #抓response

def green_elec(driver,url):
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="綠電 B.3.2.a"]')))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(1)
    page_height = driver.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );")
    # 往下滾動至頁面中間
    driver.execute_script(f"window.scrollTo(0, {page_height // 3});")
    time.sleep(1)
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
        element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"][@name="ActivityDataType"]')))
        element.click()
        element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '自動連續量測')]")))
        element.click()
        element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"][@name="EmitParaType"]')))
        element.click()
        element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '自我/量測/質能')]")))
        element.click()
    except:
        element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-status-success ant-select-single ant-select-show-arrow"][@name="ActivityDataType"]')))
        element.click()
        element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '自動連續量測')]")))
        element.click()
        element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-status-success ant-select-single ant-select-show-arrow"][@name="EmitParaType"]')))
        element.click()
        element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '自我/量測/質能')]")))
        element.click()
        time.sleep(0.5)
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '儲存')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(1)
    catch_response(driver,'輸入電力 C2(綠電)')   #抓response

def steam(driver,url,index):
    time.sleep(1)
    steam_input(driver,url,0,index)   #做蒸氣加項

def steam_input(driver,url,rfg_or_fire,index):
    personnel = ["測試人員1","測試人員2","測試人員3"]
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    element.click()
    time.sleep(1)
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//input[@placeholder="請輸入  蒸氣編號"]')))
    personnel_plate = personnel[index % len(personnel)]
    element.send_keys(personnel_plate)
    time.sleep(0.5)
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '新增')]")))
    driver.execute_script("arguments[0].click()", element)
    
    time.sleep(1)
    page_height = driver.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );")
    # 往下滾動至頁面中間
    driver.execute_script(f"window.scrollTo(0, {page_height // 3});")
    time.sleep(1)
    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '編輯')]")))
    driver.execute_script("arguments[0].click()", element)

    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-cascader ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    element.click()
    element = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'台灣')]")))
    element.click()
    element = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'2023')]")))
    element.click()
    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"][@name="ActivityDataType"]')))
    element.click()
    element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '自動連續量測')]")))
    element.click()
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