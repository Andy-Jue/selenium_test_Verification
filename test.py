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

def salience_calculate_test(driver,url):
    salience_calculate()
    # source_type = ['客戶和訪客運輸 C3','員工差旅 C3','員工通勤 C3','營運產生之廢棄物 C4','資本設備 C4','輸入能源上游排放 C4','用水(水資源管理) B.5.2.a','購買商品 C4','上游租賃資產 C4','顧問諮詢、清潔、維護等 C4','上游運輸及配送 C3','廢棄物運輸 C3','下游運輸及配送 C3','下游租賃資產 C5','產品壽命終止階段 C5','產品使用階段 C5','投資 C5', '其它間接排放 C6']
    # test_source = ['visitor','businessTravel','commuting','disposal',
    #                 'capitalGood','upstreamEmissions','waterUsage','purchasedGood','leasedAsset','consultant','upstreamTransport',
    #                 'downstreamTransport','disposalDownTransport','downLeasedAsset','downstreamDisposal','useEmission','investment','other']  

    # url = url+"/calculate/" 
    # co2_results = {} 
    # for i , j in zip(source_type,test_source):
    #     if j == 'mobileCombustion' or j == 'stationaryCombustion' or j == 'directProcessEmission' :
    #         temp = url+j
    #         driver.get(temp)
    #         time.sleep(2)
    #         page_height = driver.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );")
    #         # 往下滾動至頁面中間
    #         driver.execute_script(f"window.scrollTo(0, {page_height // 3});")
    #         time.sleep(2)
    #         print(i+ " 的 CO2排放量") 
    #         total_co2 = calculate_co2_mobileCombustion(driver)         #計算Co2
    #         logging.info(f"{i} 的 CO2排放量為 : {total_co2}")
    #         print(i+ " 的驗證 CO2排放量")
    #         co2_validate = validate_co2_mobileCombustion(driver)
    #         logging.info(f"{i} 的 CO2排放量為 : {co2_validate}")
    #         if f"{total_co2:.4f}" == f"{co2_validate:.4f}":
    #             print("CO2驗證正確!!!")
    #             logging.info("CO2驗證正確!!!")
    #         else:
    #             print("CO2驗證錯誤!!!")
    #             logging.info("CO2驗證錯誤!!!")
    #         print('\n')
    #         time.sleep(2)
    #     elif j == 'directFugitiveEmission' :
    #         temp = url+j
    #         driver.get(temp)
    #         time.sleep(2)
    #         element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    #         element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
    #         element.click()
    #         element1 = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH,'//div[@class="ant-select-item-option-content"][1]')))
    #         element1.click()
            
    #         time.sleep(2)
    #         page_height = driver.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );")
    #         # 往下滾動至頁面中間
    #         driver.execute_script(f"window.scrollTo(0, {page_height // 3});")
    #         time.sleep(2)
            
    #         print("工時計算 B.2.2.d 的 CO2排放量") 
    #         total_co2 = calculate_co2_workhour(driver)        #計算Co2
    #         logging.info(f"工時計算 B.2.2.d 的 CO2排放量為 : {total_co2}")
    #         print("工時計算 B.2.2.d 的 CO2排放量")
    #         co2_validate = validate_co2_visitor(driver)
    #         logging.info(f"工時計算 B.2.2.d 的 CO2排放量為 : {co2_validate}")
    #         if f"{total_co2:.4f}" == f"{co2_validate:.4f}" :
    #             print("CO2驗證正確!!!")
    #             logging.info("CO2驗證正確!!")
    #         else:
    #             print("CO2驗證錯誤!!!")
    #             logging.info("CO2驗證錯誤!!")
    #         print('\n')
    #         time.sleep(2)
            
    #         temp = url+j
    #         driver.get(temp)
    #         time.sleep(2)
    #         element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[text()="冷媒設備 B.2.2.d"]')))
    #         element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="冷媒設備 B.2.2.d"]')))
    #         driver.execute_script("arguments[0].click()", element)

    #         time.sleep(2)
    #         page_height = driver.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );")
    #         # 往下滾動至頁面中間
    #         driver.execute_script(f"window.scrollTo(0, {page_height // 3});")
    #         time.sleep(2)
            
    #         print("冷媒設備 B.2.2.d 的 CO2排放量") 
    #         total_co2 = calculate_co2_visitor(driver)        #計算Co2
    #         logging.info(f"冷媒設備 B.2.2.d 的 CO2排放量為 : {total_co2}")
    #         print("冷媒設備 B.2.2.d 的 CO2排放量")
    #         co2_validate = validate_co2_visitor(driver)
    #         logging.info(f"冷媒設備 B.2.2.d 的 CO2排放量為 : {co2_validate}")
    #         if f"{total_co2:.4f}" == f"{co2_validate:.4f}" :
    #             print("CO2驗證正確!!!")
    #             logging.info("CO2驗證正確!!!")
    #         else:
    #             print("CO2驗證錯誤!!!")
    #             logging.info("CO2驗證錯誤!!!")
    #         print('\n')
    #         time.sleep(2)
            
    #         temp = url+j
    #         driver.get(temp)
    #         time.sleep(2)
    #         element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[text()="消防設備 B.2.2.d"]')))
    #         element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="消防設備 B.2.2.d"]')))
    #         driver.execute_script("arguments[0].click()", element)
            
    #         time.sleep(2)
    #         page_height = driver.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );")
    #         # 往下滾動至頁面中間
    #         driver.execute_script(f"window.scrollTo(0, {page_height // 3});")
    #         time.sleep(2)
            
    #         print("消防設備 B.2.2.d 的 CO2排放量") 
    #         total_co2 = calculate_co2_visitor(driver)         #計算Co2
    #         logging.info(f"消防設備 B.2.2.d 的 CO2排放量為 : {total_co2}")
    #         print("消防設備 B.2.2.d 的 CO2排放量")
    #         co2_validate = validate_co2_visitor(driver)
    #         logging.info(f"消防設備 B.2.2.d 的 CO2排放量為 : {co2_validate}")
    #         if f"{total_co2:.4f}" == f"{co2_validate:.4f}" :
    #             print("CO2驗證正確!!!")
    #             logging.info("CO2驗證正確!!!")
    #         else:
    #             print("CO2驗證錯誤!!!")
    #             logging.info("CO2驗證錯誤!!!")
    #         print('\n')
    #         time.sleep(2)
            
    #     elif j == 'electricity' :
    #         temp = url+j
    #         driver.get(temp)
            
    #         time.sleep(2)
    #         page_height = driver.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );")
    #         # 往下滾動至頁面中間
    #         driver.execute_script(f"window.scrollTo(0, {page_height // 3});")
    #         time.sleep(2)
            
    #         print("一般用電 B.3.2.a 的 CO2排放量") 
    #         total_co2 = calculate_co2_workhour(driver)         #計算Co2
    #         logging.info(f"一般用電 B.3.2.a 的 CO2排放量為 : {total_co2}")
    #         print("一般用電 B.3.2.a 的驗證 CO2排放量")
    #         co2_validate = validate_co2_workhour(driver)
    #         logging.info(f"一般用電 B.3.2.a 的 CO2排放量為 : {co2_validate}")
    #         co2_results["一般用電 B.3.2.a"] = co2_validate
    #         if f"{total_co2:.4f}" == f"{co2_validate:.4f}" :
    #             print("CO2驗證正確!!!")
    #             logging.info("CO2驗證正確!!!")
    #         else:
    #             print("CO2驗證錯誤!!!")
    #             logging.info("CO2驗證錯誤!!!")
    #         print('\n')
    #         time.sleep(2)

    #     elif j == 'upstreamEmissions' :
    #         temp = url+j
    #         driver.get(temp)
    #         time.sleep(2)
    #         element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[text()="輸入電力上游 B.5.2.a"]')))
    #         element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="輸入電力上游 B.5.2.a"]')))
    #         driver.execute_script("arguments[0].click()", element)
            
    #         time.sleep(2)
    #         page_height = driver.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );")
    #         # 往下滾動至頁面中間
    #         driver.execute_script(f"window.scrollTo(0, {page_height // 3});")
    #         time.sleep(2)
            
    #         print("輸入電力上游 B.5.2.a 的 CO2排放量") 
    #         total_co2 = calculate_co2_visitor(driver)          #計算Co2
    #         logging.info(f"輸入電力上游 B.5.2.a 的 CO2排放量為 : {total_co2}")
    #         print("輸入電力上游 B.5.2.a 的驗證 CO2排放量")
    #         co2_validate = validate_co2_visitor(driver)
    #         logging.info(f"輸入電力上游 B.5.2.a 的驗證 CO2排放量為 : {co2_validate}")
    #         co2_results["輸入電力上游 B.5.2.a"] = co2_validate
    #         if f"{total_co2:.4f}" == f"{co2_validate:.4f}" :
    #             print("CO2驗證正確!!!")
    #             logging.info("CO2驗證正確!!!")
    #         else:
    #             print("CO2驗證錯誤!!!")
    #             logging.info("CO2驗證錯誤!!!")
    #         print('\n')
    #         time.sleep(2)
            
    #         temp = url+j
    #         driver.get(temp)
    #         time.sleep(2)
    #         element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[text()="其他輸入能源上游 B.5.2.a"]')))
    #         element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="其他輸入能源上游 B.5.2.a"]')))
    #         driver.execute_script("arguments[0].click()", element)
            
    #         time.sleep(2)
    #         page_height = driver.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );")
    #         # 往下滾動至頁面中間
    #         driver.execute_script(f"window.scrollTo(0, {page_height // 3});")
    #         time.sleep(2)
            
    #         print("其他輸入能源上游 B.5.2.a 的 CO2排放量") 
    #         total_co2 = calculate_co2_visitor(driver)          #計算Co2
    #         logging.info(f"其他輸入能源上游 B.5.2.a 的 CO2排放量為 : {total_co2}")
    #         print("其他輸入能源上游 B.5.2.a 的驗證 CO2排放量")
    #         co2_validate = validate_co2_visitor(driver)
    #         logging.info(f"其他輸入能源上游 B.5.2.a 的 CO2排放量為 : {co2_validate}")
    #         co2_results["其他輸入能源上游 B.5.2.a"] = co2_validate
    #         if f"{total_co2:.4f}" == f"{co2_validate:.4f}" :
    #             print("CO2驗證正確!!!")
    #             logging.info("CO2驗證正確!!!")
    #         else:
    #             print("CO2驗證錯誤!!!")
    #             logging.info("CO2驗證錯誤!!!")
    #         print('\n')
    #         time.sleep(2)

        
    #     else:
    #         temp = url+j
    #         driver.get(temp)
            
    #         time.sleep(2)
    #         page_height = driver.execute_script("return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );")
    #         # 往下滾動至頁面中間
    #         driver.execute_script(f"window.scrollTo(0, {page_height // 3});")
    #         time.sleep(2)
            
    #         print(i+ " 的 CO2排放量") 
    #         total_co2 = calculate_co2_visitor(driver)         #計算Co2
    #         logging.info(f"{i} 的 CO2排放量為 : {total_co2}")
    #         print(i+ " 的驗證 CO2排放量")
    #         co2_validate = validate_co2_visitor(driver)
    #         logging.info(f"{i} 的 CO2排放量為 : {co2_validate}")
            
    #         co2_results[i] = co2_validate 
            
    #         if f"{total_co2:.4f}" == f"{co2_validate:.4f}" :
    #             print("CO2驗證正確!!!")
    #             logging.info("CO2驗證正確!!!")
    #         else:
    #             print("CO2驗證錯誤!!!")
    #             logging.info("CO2驗證錯誤!!!")
    #         print('\n')
    #         time.sleep(2)
    # print(co2_results)
    # co2_combined_emissions = {}

    # for key, value in co2_results.items():
    #     # 去除括號內的內容，僅保留合併基礎部分
    #     if "一般用電" in key:
    #         base_key = "輸入電力"
    #     elif "客戶和訪客運輸" in key :
    #         base_key = "運輸客戶與訪客"
    #     elif "員工差旅" in key :
    #         base_key = "業務旅運"
    #     elif "員工通勤" in key :
    #         base_key = "員工通勤"
    #     elif "輸入電力上游" in key:
    #         base_key = "輸入電力上游"
    #     elif "其他輸入能源上游" in key:
    #         base_key = "輸入能源上游"
    #     elif "用水(水資源管理)" in key:
    #         base_key = "用水"
    #     elif "上游運輸及配送" in key:
    #         base_key = "上游運輸"
    #     elif "廢棄物運輸" in key:
    #         base_key = "廢棄物運送"
    #     elif "下游運輸及配送" in key:
    #         base_key = "下游運輸"
    #     elif "C" in key :
    #         base_key = key.split("C")[0]
    #     else:
    #         base_key = key
        
        
    #     # 如果該基礎鍵已存在於新字典中，則累加數值
    #     if base_key in co2_combined_emissions:
    #         co2_combined_emissions[base_key] = round(co2_combined_emissions[base_key] + round(float(value) / 1000, 10), 10)
    #     else:
    #         # 如果該基礎鍵不存在，則初始化
    #         co2_combined_emissions[base_key] = round((float(value) / 1000), 10)
    # print(co2_combined_emissions)
    # # 輸出合併後的結果
    # for key, total in co2_combined_emissions.items():
        
    #     print(f"{key.strip()}: {total :.10f}")        
    
    # time.sleep(1)
    # salience_calculate_results = salience_calculate(driver,url)

    # combined_emissions = {}

    # for key, value in salience_calculate_results.items():
    #     # 去除括號內的內容，僅保留合併基礎部分
    #     if "輸入電力上游" in key:
    #         base_key = "輸入電力上游"
    #     elif "輸入能源上游" in key :
    #         base_key = "輸入能源上游"
    #     elif "上游運輸" in key :
    #         base_key = "上游運輸"
    #     elif "廢棄物運送" in key :
    #         base_key = "廢棄物運送"
    #     elif "下游運輸" in key :
    #         base_key = "下游運輸"
    #     elif "(" in key :
    #         base_key = key.split('(')[0]
    #     elif "C" in key :
    #         base_key = key.split('C')[0]
    #     else:
    #         base_key = key
        
        
    #     # 如果該基礎鍵已存在於新字典中，則累加數值
    #     if base_key in combined_emissions:
    #         combined_emissions[base_key] = round(combined_emissions[base_key] + round(float(value), 10), 10)
    #     else:
    #         # 如果該基礎鍵不存在，則初始化
    #         combined_emissions[base_key] = round(float(value), 10)
    # print(combined_emissions)
    
    # # 輸出合併後的結果
    # for key, total in combined_emissions.items():
    #     print(f"{key.strip()}: {total :.10f}")
    # time.sleep(1)
    # for key in combined_emissions:
    #     if key in co2_combined_emissions:
    #         if round(combined_emissions[key],10) == round(co2_combined_emissions[key],10):
    #             print(f"{key}的值: {combined_emissions[key]} 驗證正確")
    #         else:
    #             print(f"{key}的值: {combined_emissions[key]} 驗證錯誤 (正確值為: {co2_combined_emissions[key]})")
    #     else:
    #         print(f"{key}不在計算當中")
    
        
def calculate_co2_mobileCombustion(driver):
    total_carbon_emissions = 0
    
    has_next_page = True
    
    while has_next_page :
        
        table = driver.find_element(By.TAG_NAME, 'table')
    
        # 定位表頭行元素
        thead_row = table.find_element(By.XPATH, './/thead/tr')
    
        # 找到 "燃料種類" 和 "碳排放量" 在表頭中的索引位置
        carbon_emissions_index = -1
        header_cells = thead_row.find_elements(By.TAG_NAME, 'th')
        for i, cell in enumerate(header_cells):
            if cell.text == '碳排放量(kgCO2e)':
                carbon_emissions_index = i
    
    
        # 檢查是否找到 "燃料種類" 和 "碳排放量" 的索引位置
        if carbon_emissions_index != -1:
            # 定位所有的 <tbody> 元素
            tbodies = table.find_elements(By.TAG_NAME, 'tbody')
    
            # 遍歷每個 <tbody>
            for tbody in tbodies:
                # 定位該 <tbody> 中所有的行元素
                rows = tbody.find_elements(By.TAG_NAME, 'tr')
    
                # 遍歷每個行元素
                for row in rows[1:]:
                
                
                # 定位燃料種類和碳排放量元素
                    carbon_emissions_elements = row.find_elements(By.XPATH, f'.//td[{carbon_emissions_index + 1}]')
    
                # 獲取燃料種類和碳排放量的值
                    for carbon_emissions_element in carbon_emissions_elements:
                        carbon_emissions = float(carbon_emissions_element.text)
                        
                        total_carbon_emissions += carbon_emissions
            
                # 更新燃料種類的碳排放量總和
            try:
                next_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//li[@title='下一頁']//button[@class='ant-pagination-item-link']")))
                ActionChains(driver).move_to_element(next_button).click().perform()
                
                time.sleep(2)
                    
            except Exception as e:
                # print(f"An error occurred: {e}")
                print(f'碳排放量總和：{total_carbon_emissions}')
                has_next_page = False

        else:
            print('未找到 "碳排放量"')
            
    return total_carbon_emissions

def validate_co2_mobileCombustion(driver):
    table = driver.find_element(By.XPATH, '//table[contains(@class, "table table-sm table-bordered text-end")]')
    thead_row = table.find_element(By.XPATH, './/thead/tr')

    # 找到 "排放源" 和 "排放量 kgCO2e" 在表頭中的索引位置
    emissions_index = -1
    header_cells = thead_row.find_elements(By.TAG_NAME, 'th')
    for i, cell in enumerate(header_cells):
        
        if cell.text == '排放量 KgCO2e':
            emissions_index = i
            
    # 檢查是否找到 "排放源" 和 "排放量 kgCO2e" 的索引位置
    if emissions_index != -1:
        # 定位 <tbody> 元素
        tbody = table.find_element(By.TAG_NAME, 'tbody')

        # 定位所有的行元素
        rows = tbody.find_elements(By.TAG_NAME, 'tr')

        # 創建一個字典來跟踪每個排放源的排放量總和
        # emissions_dict = {}
        total_emissions_elements = 0

        # 遍歷每個行元素
        row_end = rows[-1]
        
        emissions_elements = row_end.find_element(By.XPATH, f'.//td[{emissions_index}]')
        emissions_element = float(emissions_elements.text)
            
        total_emissions_elements += emissions_element


        # 更新排放源的排放量總和
        print(f'排放量 KgCO2e : {total_emissions_elements}')
    else:
        print('未找到標題為 "排放量 KgCO2e" 的列')

    return total_emissions_elements


def calculate_co2_otherCompound(driver):
    total_otherCompoundn_emissions = 0
    
    has_next_page = True
    
    while has_next_page :
        table = driver.find_element(By.TAG_NAME, 'table')
    
        # 定位表頭行元素
        thead_row = table.find_element(By.XPATH, './/thead/tr')
    
        # 找到 "燃料種類" 和 "碳排放量" 在表頭中的索引位置
        carbon_emissions_index = -1
        header_cells = thead_row.find_elements(By.TAG_NAME, 'th')
        for i, cell in enumerate(header_cells):
            if cell.text == '碳排放量(kgCO2e)':
                carbon_emissions_index = i
    
    
        # 檢查是否找到 "燃料種類" 和 "碳排放量" 的索引位置
        if carbon_emissions_index != -1:
            # 定位所有的 <tbody> 元素
            tbodies = table.find_elements(By.TAG_NAME, 'tbody')
    
            # 創建一個字典來跟踪每種燃料種類的碳排放量總和
            total_otherCompoundn_emissions = 0
    
            # 遍歷每個 <tbody>
            for tbody in tbodies:
                # 定位該 <tbody> 中所有的行元素
                rows = tbody.find_elements(By.TAG_NAME, 'tr')
    
                # 遍歷每個行元素
                for row in rows[1:]:
                
                
                # 定位燃料種類和碳排放量元素
                    carbon_emissions_elements = row.find_elements(By.XPATH, f'.//td[{carbon_emissions_index + 1}]')
    
                # 獲取燃料種類和碳排放量的值
                    for carbon_emissions_element in carbon_emissions_elements:
                        carbon_emissions = float(carbon_emissions_element.text)
                        
                        total_otherCompoundn_emissions += carbon_emissions
    
                # 更新燃料種類的碳排放量總和
            try:
                next_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//li[@title='下一頁']//button[@class='ant-pagination-item-link']")))
                ActionChains(driver).move_to_element(next_button).click().perform()
                
                time.sleep(2)
                
              
                    
            except Exception as e:
                # print(f"An error occurred: {e}")
                print(f'碳排放量總和：{total_otherCompoundn_emissions}')
                has_next_page = False
    
            # 輸出每種燃料種類的碳排放量總和    
    
            # 輸出每種燃料種類的碳排放量總和
        else:
            print('未找到 "碳排放量"')

    return total_otherCompoundn_emissions

def validate_co2_otherCompound(driver):
    table = driver.find_element(By.XPATH, '//table[contains(@class, "table table-sm table-bordered text-end")]')
    thead_row = table.find_element(By.XPATH, './/thead/tr')

    # 找到 "排放源" 和 "排放量 kgCO2e" 在表頭中的索引位置
    emissions_index = -1
    header_cells = thead_row.find_elements(By.TAG_NAME, 'th')
    for i, cell in enumerate(header_cells):
        
        if cell.text == '排放量 KgCO2e':
            emissions_index = i
            

           
        
    # 檢查是否找到 "排放源" 和 "排放量 kgCO2e" 的索引位置
    if emissions_index != -1:
        # 定位 <tbody> 元素
        tbody = table.find_element(By.TAG_NAME, 'tbody')

        # 定位所有的行元素
        rows = tbody.find_elements(By.TAG_NAME, 'tr')

        # 創建一個字典來跟踪每個排放源的排放量總和
        # emissions_dict = {}
        total_otherCompound_elements = 0

        # 遍歷每個行元素
        row_end = rows[-1]
        
        emissions_elements = row_end.find_element(By.XPATH, f'.//td[{emissions_index + 1}]')
        emissions_element = float(emissions_elements.text)
            
        total_otherCompound_elements += emissions_element

            
             

            # 獲取排放源和排放量的值
        # emissions = float(emissions_element.text)
        
        

        # 更新排放源的排放量總和
        print(f'排放量 KgCO2e : {total_otherCompound_elements}')
    else:
        print('未找到標題為 "排放量 KgCO2e" 的列')

    return total_otherCompound_elements


def calculate_co2_visitor(driver):
    total_visitor_emissions = 0
    
    has_next_page = True
    
    while has_next_page :
        table = driver.find_element(By.TAG_NAME, 'table')
    
        # 定位表頭行元素
        thead_row = table.find_element(By.XPATH, './/thead/tr')
    
        # 找到 "燃料種類" 和 "碳排放量" 在表頭中的索引位置
        carbon_emissions_index = -1
        header_cells = thead_row.find_elements(By.TAG_NAME, 'th')
        for i, cell in enumerate(header_cells):
            if cell.text == '碳排放量(kgCO2e)':
                carbon_emissions_index = i
    
    
        # 檢查是否找到 "燃料種類" 和 "碳排放量" 的索引位置
        if carbon_emissions_index != -1:
            # 定位所有的 <tbody> 元素
            tbodies = table.find_elements(By.TAG_NAME, 'tbody')
    
            # 創建一個字典來跟踪每種燃料種類的碳排放量總和
            
    
            # 遍歷每個 <tbody>
            for tbody in tbodies:
                # 定位該 <tbody> 中所有的行元素
                rows = tbody.find_elements(By.TAG_NAME, 'tr')
    
                # 遍歷每個行元素
                for row in rows[1:]:
                
                
                # 定位燃料種類和碳排放量元素
                    carbon_emissions_elements = row.find_elements(By.XPATH, f'.//td[{carbon_emissions_index + 1}]')
    
                # 獲取燃料種類和碳排放量的值
                    for carbon_emissions_element in carbon_emissions_elements:
                        carbon_emissions = float(carbon_emissions_element.text)
                        
                        total_visitor_emissions += carbon_emissions
    
                # 更新燃料種類的碳排放量總和
            try:
                next_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//li[@title='下一頁']//button[@class='ant-pagination-item-link']")))
                ActionChains(driver).move_to_element(next_button).click().perform()
                
                time.sleep(2)
                
            except Exception :
                print(f'碳排放量總和：{total_visitor_emissions}')
                has_next_page = False    
    
            # 輸出每種燃料種類的碳排放量總和
        else:
            print('未找到 "碳排放量"')
            break

    return total_visitor_emissions

def validate_co2_visitor(driver):
    table = driver.find_element(By.XPATH, '//table[contains(@class, "table table-sm table-bordered text-end")]')
    thead_row = table.find_element(By.XPATH, './/thead/tr')

    # 找到 "排放源" 和 "排放量 kgCO2e" 在表頭中的索引位置
    emissions_index = -1
    header_cells = thead_row.find_elements(By.TAG_NAME, 'th')
    for i, cell in enumerate(header_cells):
        
        if cell.text == '排放量 KgCO2e':
            emissions_index = i
           

           
        
    # 檢查是否找到 "排放源" 和 "排放量 kgCO2e" 的索引位置
    if emissions_index != -1:
        # 定位 <tbody> 元素
        tbody = table.find_element(By.TAG_NAME, 'tbody')

        # 定位所有的行元素
        rows = tbody.find_elements(By.TAG_NAME, 'tr')

        # 創建一個字典來跟踪每個排放源的排放量總和
        # emissions_dict = {}
        total_emissions_elements = 0

        # 遍歷每個行元素
        row_end = rows[-1]
        
        emissions_elements = row_end.find_element(By.XPATH, f'.//td[{emissions_index + 1}]')
        emissions_element = float(emissions_elements.text)
            
        total_emissions_elements += emissions_element

            
             

            # 獲取排放源和排放量的值
        # emissions = float(emissions_element.text)
        
        

        # 更新排放源的排放量總和
        print(f'排放量 KgCO2e : {total_emissions_elements}')
    else:
        print('未找到標題為 "排放量 KgCO2e" 的列')

    return total_emissions_elements


def calculate_co2_workhour(driver):
    total_visitor_emissions = 0
        
        
        
    table = driver.find_element(By.TAG_NAME, 'table')

    # 定位表頭行元素
    tbody_row = table.find_element(By.TAG_NAME, 'tbody')
    rows = tbody_row.find_elements(By.TAG_NAME, 'tr')
    carbon_emissions_index = None
    for row in rows:
        
        header_cells = row.find_elements(By.XPATH, './/td')
        for cell in header_cells:
            if cell.text == '碳排放量(kgCO2e)':
                carbon_emissions_index = row
                
  
    if carbon_emissions_index is not None:
        # 定位該 <tbody> 中所有的行元素
        carbon_emissions_elements = carbon_emissions_index.find_elements(By.XPATH, './/td')
        
        # 獲取燃料種類和碳排放量的值
        for carbon_emissions_element in carbon_emissions_elements[1:]:
            carbon_emissions = float(carbon_emissions_element.text)
                
            total_visitor_emissions += carbon_emissions

        # 更新燃料種類的碳排放量總和

        print(f'碳排放量總和：{total_visitor_emissions}')
              

        # 輸出每種燃料種類的碳排放量總和
    else:
        print('未找到 "KgCO2e"')
       
    
    return total_visitor_emissions
def validate_co2_workhour(driver):
    table = driver.find_element(By.XPATH, '//table[contains(@class, "table table-sm table-bordered text-end")]')
    thead_row = table.find_element(By.XPATH, './/thead/tr')

    # 找到 "排放源" 和 "排放量 kgCO2e" 在表頭中的索引位置
    emissions_index = -1
    header_cells = thead_row.find_elements(By.TAG_NAME, 'th')
    for i, cell in enumerate(header_cells):
        
        if cell.text == '排放量 KgCO2e':
            emissions_index = i
            

           
        
    # 檢查是否找到 "排放源" 和 "排放量 kgCO2e" 的索引位置
    if emissions_index != -1:
        # 定位 <tbody> 元素
        tbody = table.find_element(By.TAG_NAME, 'tbody')

        # 定位所有的行元素
        rows = tbody.find_elements(By.TAG_NAME, 'tr')

        # 創建一個字典來跟踪每個排放源的排放量總和
        # emissions_dict = {}
        total_emissions_elements = 0

        # 遍歷每個行元素
        row_end = rows[-1]        
        emissions_elements = row_end.find_element(By.XPATH, f'.//td[{emissions_index} +1]')
        emissions_element = float(emissions_elements.text)
            
        total_emissions_elements += emissions_element

            # 獲取排放源和排放量的值
        # emissions = float(emissions_element.text)
        
        

        # 更新排放源的排放量總和
        print(f'排放量 KgCO2e : {total_emissions_elements}')
    else:
        print('未找到標題為 "排放量 KgCO2e" 的列')

    return total_emissions_elements



def salience_calculate(driver,url):
    sign(driver,url)
    time.sleep(2)
    has_next_page = True
    emissions_dict = {}
    while has_next_page:
        table = driver.find_element(By.XPATH, '//div[@class = "ant-table-wrapper css-dts6b9"]')
    
        # 定位 "生命週期階段" 和 "碳排放量 (噸) 總量" 在表頭中的索引位置
        score_1_index = 5
        score_2_index = 6
        score_3_index = 7
        score_4_index = 8
        score_5_index = 9
        score_6_index = 10
        source_index = 2
    
        # 定位 <tbody> 元素
        tbody = table.find_element(By.TAG_NAME, 'tbody')
    
        # 定位所有的行元素
        rows = tbody.find_elements(By.TAG_NAME, 'tr')
    
        # 創建一個字典來跟踪每個排放源的排放量總和
        
        for row in rows[1:]:
            #定位生命週期表正確欄位
            source_element = row.find_element(By.XPATH, f'.//td[{source_index + 1}]')
            score_1_element = row.find_element(By.XPATH, f'.//td[{score_1_index + 1}]')
            score_2_element = row.find_element(By.XPATH, f'.//td[{score_2_index + 1}]')
            score_3_element = row.find_element(By.XPATH, f'.//td[{score_3_index + 1}]')
            score_4_element = row.find_element(By.XPATH, f'.//td[{score_4_index + 1}]')
            score_5_element = row.find_element(By.XPATH, f'.//td[{score_5_index + 1}]')
            score_6_element = row.find_element(By.XPATH, f'.//td[{score_6_index + 1}]')
            
            # 轉換為正常的格式
            source_text = source_element.text.strip()
            score_1_value = score_1_element.text.strip()
            score_2_value = score_2_element.text.strip()
            score_3_value = score_3_element.text.strip()
            score_4_value = score_4_element.text.strip()
            score_5_value = score_5_element.text.strip()
            score_6_value = score_6_element.text.strip()
            # print(source_text)
            # print(emissions_value)
            score_values = [score_1_value, score_2_value, score_3_value, score_4_value, score_5_value, score_6_value]
            if source_text in score_1_value:
                emissions_dict[source_text].append(score_values)
            else:
                emissions_dict[source_text] = score_values
        try:
            next_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//li[@title='下一頁']//button[@class='ant-pagination-item-link']")))
            ActionChains(driver).move_to_element(next_button).click().perform()
            
            time.sleep(2)
            
        except Exception :
            has_next_page = False
    print(emissions_dict)
    return emissions_dict

def sign(driver,url):
    driver.get("https://220.132.206.5:8888/")
    time.sleep(2)
    wait = WebDriverWait(driver, 10)
    autotest_td = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//td[text()='測試盤查']")))

# 找到父 <tr> 元素
    parent_tr = autotest_td.find_element(By.XPATH, './parent::tr')

    button = parent_tr.find_element(By.XPATH, './/button[@type="button" and contains(@class, "ant-btn ant-btn-round ant-btn-link ant-btn-icon-only text-primary  border-0")]')
#JavaScript寫法
    driver.execute_script("arguments[0].click()", button)
    time.sleep(1)
    
    link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "顯著性分析")))   #點擊排放資料輸入
    link.click()
    