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

def excel_output(driver,url,temp_func):
    source_type = ['移動源 C1', '固定燃燒源 C1', '工業製程 C1', '人為逸散 C1', '其他關注類物質 C1', '輸入電力 C2',
                   '客戶和訪客運輸 C3', '員工差旅 C3', '員工通勤 C3', '營運產生之廢棄物 C4', '資本設備 C4',
                   '輸入能源上游排放 C4', '用水(水資源管理) B.5.2.a', '購買商品 C4', '上游租賃資產 C4',
                   '顧問諮詢、清潔、維護等 C4', '上游運輸及配送 C3', '廢棄物運輸 C3', '下游運輸及配送 C3',
                   '下游租賃資產 C5', '產品壽命終止階段 C5', '產品使用階段 C5', '投資 C5', '其它間接排放 C6']
    for i in source_type:
        if i == '人為逸散 C1':
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))
            link.click() 
            time.sleep(2)
            output_func(driver,source_type,temp_func,i,1,0,0,0)

            element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="冷媒設備 B.2.2.d"]')))
            driver.execute_script("arguments[0].click()", element)
            time.sleep(2)
            output_func(driver,source_type,temp_func,i,2,0,0,0)

            element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="消防設備 B.2.2.d"]')))
            driver.execute_script("arguments[0].click()", element)
            time.sleep(2)
            output_func(driver,source_type,temp_func,i,3,0,0,0)
        
        elif i == '輸入電力 C2':
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))
            link.click() 
            time.sleep(2)
            output_func(driver,source_type,temp_func,i,0,1,0,0)
            
            element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[text()="綠電 B.3.2.a"]')))
            element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="綠電 B.3.2.a"]')))
            driver.execute_script("arguments[0].click()", element)
            output_func(driver,source_type,temp_func,i,0,2,0,0)

        elif i == '輸入蒸汽 C2':
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))
            link.click() 
            time.sleep(2)
            output_func(driver,source_type,temp_func,i,0,0,1,0)

            element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="蒸氣減項 B.3.2.b"]')))
            driver.execute_script("arguments[0].click()", element)
            output_func(driver,source_type,temp_func,i,0,0,2,0)
        elif i =='輸入能源上游排放 C4':
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))
            link.click() 
            time.sleep(2)
            output_func(driver,source_type,temp_func,i,0,0,0,1)

            element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="其他輸入能源上游 B.5.2.a"]')))
            driver.execute_script("arguments[0].click()", element)
            output_func(driver,source_type,temp_func,i,0,0,0,2)
            time.sleep(2)
        else:
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))
            link.click()
            time.sleep(2)
            output_func(driver, source_type, temp_func, i, 0, 0, 0, 0)
            time.sleep(2)


def output_func(driver, source_type, temp_func, i, rfg_or_fire, elec_or_green, plus_or_minus, ups_or_orther):
    # 點擊 "匯入 & 匯出" 按鈕
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '匯入 & 匯出')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(2)

    # 選擇項目
    select = [
        [rfg_or_fire, '人為逸散 C1(工時計算)', '人為逸散 C1(冷媒設備)', '人為逸散 C1(消防設備)'],
        [elec_or_green, '輸入電力 C2(一般用電)', '輸入電力 C2(綠電)'],
        [plus_or_minus, '輸入蒸汽 C2(蒸汽加項)', '輸入蒸汽 C2(蒸汽減項)'],
        [ups_or_orther, '輸入電力上游 B.5.2a', '其他輸入能源上游 B.5.2.a']
    ]

    this_time_select = None
    for each in select:
        if 0 < each[0] < len(each):  # 確保 index 在合理範圍內
            this_time_select = each[each[0]]
            break  # 找到一個就退出循環，如果是這樣的邏輯
    # 執行匯出操作
    if temp_func == 'output':  # 匯出 excel 資料
        import_element = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//ul//button[span[text()="匯出"]]')))
        actions = ActionChains(driver)
        actions.move_to_element(import_element).click().perform()
        if i != '人為逸散 C1' and i != '輸入電力 C2' and i != '輸入蒸汽 C2' and i != '輸入能源上游排放 C4':
            catch_response(driver, i)
            time.sleep(3)
        else:
            catch_response(driver, this_time_select)
            time.sleep(3)


    elif temp_func =='download':             #單純模板下載
        import_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//ul//button[span[text()="模板下載"]]')))
        driver.execute_script("arguments[0].click()", import_element)
        time.sleep(3)
        # 模板下載沒有response