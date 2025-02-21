import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains


def significance_edit(driver,url,site_num):
    sign_Significance(driver,url,site_num)
    time.sleep(1)
    has_next_page = True
    while has_next_page:
        edit_buttons = driver.find_elements(By.XPATH, '//button[contains(text(), "編輯")]')
        for button in edit_buttons:
            time.sleep(5)
            driver.execute_script("arguments[0].click();", button)
            time.sleep(1)
            input_data_C1_C2(driver,url)
            
        try:
            next_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//li[@title='下一頁']//button[@class='ant-pagination-item-link']")))
            ActionChains(driver).move_to_element(next_button).click().perform()
            
            time.sleep(2)
            
        except Exception :
            has_next_page = False
            print("顯著性評分最後一頁完成")
    time.sleep(1)
    sign_sources(driver,url,site_num)
    time.sleep(1)
def input_data_C1_C2(driver,url):
    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME, "ant-modal-body")))

    form_element = driver.find_element(By.CLASS_NAME, "ant-modal-body")      

    input_elements = form_element.find_elements(By.TAG_NAME, "input")        

    new_inputelements =[]

    for i in input_elements:                     
        if i.get_attribute("disabled") != None and i.get_attribute("type")!="search":
            continue
        else:
            new_inputelements.append(i)
    for input_element in new_inputelements:
    
        if input_element.get_attribute("type")=="search":  
            if input_element.get_attribute("id") in ["significant_3008","significant_3048","significant_101","significant_2521","significant_2241"]:
                ActionChains(driver).move_to_element(input_element).click().perform()
                time.sleep(1)
                element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[@class='ant-select-item-option-content' and text()='明確要求揭露']")))
                element.click()
                time.sleep(1)
            elif input_element.get_attribute("id") in ["significant_3009","significant_3049","significant_102","significant_2522","significant_2242"]:
                ActionChains(driver).move_to_element(input_element).click().perform()
                time.sleep(1)
                element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "(//div[@class='ant-select-item-option-content' and text()='可能未來要求揭露'])[2]")))
                element.click()
                time.sleep(1)
            elif input_element.get_attribute("id") in ["significant_3010","significant_3050","significant_103","significant_2523","significant_2243"]:
                ActionChains(driver).move_to_element(input_element).click().perform()
                element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[@class='ant-select-item-option-content' and text()='高階分析單一排放量大於10%']")))
                element.click()
                time.sleep(1)
            elif input_element.get_attribute("id") in ["significant_3011","significant_3051","significant_104","significant_2524","significant_2244"]:
                ActionChains(driver).move_to_element(input_element).click().perform()
                element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[@class='ant-select-item-option-content' and text()='控制力較佳能取得特定廠址數據、一級數據']")))
                element.click()
                time.sleep(1)
            elif input_element.get_attribute("id") in ["significant_3012","significant_3052","significant_105","significant_2525","significant_2245"]:
                ActionChains(driver).move_to_element(input_element).click().perform()
                element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[@class='ant-select-item-option-content' and text()='揭露後能鼓勵員工降低排碳量並有助於推動低碳生活']"))) #移動源係數
                element.click()
                time.sleep(1)
    try:
        element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '確 定')]")))
        driver.execute_script("arguments[0].click()", element)
        time.sleep(2)
        print("編輯評分成功")
    except:
        print("顯著性編輯失敗")

def sign_Significance(driver,url,site_num):
    site_dict = {1: "https://prod.netzero.com.tw", 2: "https://demo2.netzero.com.tw",
                 3: "https://14064-1.com", 4: "https://ghg.netzeroyun.com",
                 5: "https://220.132.206.5:666", 6: "https://220.132.206.5:8888",
                 7: "https://220.132.206.5:9999"}
    url = site_dict[site_num]
    driver.get(url)
    time.sleep(2)
    wait = WebDriverWait(driver, 10)
    if site_num == 3:
        autotest_td = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//td[text()='亞O源']"))  # 要測試的盤查計畫名稱
        )
    elif site_num == 5:
        autotest_td = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//td[text()='A廠區測試']"))  # 要測試的盤查計畫名稱
        )
    else:
        autotest_td = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//td[text()='測試盤查']"))  # 要測試的盤查計畫名稱
        )

# 找到父 <tr> 元素
    parent_tr = autotest_td.find_element(By.XPATH, './parent::tr')

    button = parent_tr.find_element(By.XPATH, './/button[@type="button" and contains(@class, "ant-btn ant-btn-round ant-btn-link ant-btn-icon-only text-primary  border-0")]')
#JavaScript寫法
    driver.execute_script("arguments[0].click()", button)
    time.sleep(1)
    
    link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "顯著性分析")))   #點擊排放資料輸入
    link.click()
def sign_sources(driver,url,site_num):
    site_dict = {1: "https://prod.netzero.com.tw", 2: "https://demo2.netzero.com.tw",
                 3: "https://14064-1.com", 4: "https://ghg.netzeroyun.com",
                 5: "https://220.132.206.5:666", 6: "https://220.132.206.5:8888",
                 7: "https://220.132.206.5:9999"}
    url = site_dict[site_num]
    driver.get(url)
    time.sleep(2)
    wait = WebDriverWait(driver, 10)
    if site_num == 3:
        autotest_td = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//td[text()='亞O源']"))  # 要測試的盤查計畫名稱
        )
    elif site_num == 5:
        autotest_td = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//td[text()='A廠區測試']"))  # 要測試的盤查計畫名稱
        )
    else:
        autotest_td = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//td[text()='測試盤查']"))  # 要測試的盤查計畫名稱
        )

# 找到父 <tr> 元素
    parent_tr = autotest_td.find_element(By.XPATH, './parent::tr')

    button = parent_tr.find_element(By.XPATH, './/button[@type="button" and contains(@class, "ant-btn ant-btn-round ant-btn-link ant-btn-icon-only text-primary  border-0")]')
#JavaScript寫法
    driver.execute_script("arguments[0].click()", button)
    time.sleep(1)
    
    link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "排放資料輸入")))   #點擊排放資料輸入
    link.click()
