from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from response import catch_response
import os
import time
import sys
def resource_path(relative_path):
   #取得資源檔案的絕對路徑，無論是開發環境或打包後的環境
    try:
        # PyInstaller 建立臨時資料夾，並將路徑儲存在 _MEIPASS 中
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def excel_input(driver, url):
    source_to_file_map = {
        '移動源 C1': '(新增)正確數值_Vehicle.xlsx',
        '固定燃燒源 C1': '(新增)正確數值_StationaryCombustionSource.xlsx',
        '工業製程 C1': '(新增)正確數值_IndustrialProcess.xlsx',
        '人為逸散 C1': '(新增)正確數值_workhour.xlsx',
        '其他關注類物質 C1': '(新增)正確數值_other.xlsx',
        '輸入電力 C2': '(新增)正確數值_Elec.xlsx',
        '客戶和訪客運輸 C3': '(新增)正確數值_運輸客戶與訪客.xlsx',
        '員工差旅 C3': '(新增)正確數值_業務旅運(商務旅行).xlsx',
        '員工通勤 C3': '(新增)正確數值_員工通勤.xlsx',
        '營運產生之廢棄物 C4': '(新增)正確數值_廢棄物處理.xlsx',
        '資本設備 C4': '(新增)正確數值_資本財(產生之排放).xlsx',
        '輸入能源上游排放 C4': '(新增)正確數值_輸入電力上游.xlsx',
        '用水(水資源管理) B.5.2.a': '(新增)正確數值_用水(水資源管理).xlsx',
        '購買商品 C4': '(新增)正確數值_購買商品.xlsx',
        '上游租賃資產 C4': '(新增)正確數值_租賃設備資產使用.xlsx',
        '顧問諮詢、清潔、維護等 C4': '(新增)正確數值_組織使用之服務.xlsx',
        '上游運輸及配送 C3': '(新增)正確數值_上游運輸.xlsx',
        '廢棄物運輸 C3': '(新增)正確數值_廢棄物運輸.xlsx',
        '下游運輸及配送 C3': '(新增)正確數值_下游運輸.xlsx',
        '下游租賃資產 C5': '(新增)正確數值_下游承租資產.xlsx',
        '產品壽命終止階段 C5': '(新增)正確數值_產品生命終止階段.xlsx',
        '產品使用階段 C5': '(新增)正確數值_產品使用階段.xlsx',
        '投資 C5': '(新增)正確數值_投資.xlsx',
        '其它間接排放 C6': '(新增)正確數值_其它.xlsx'
    }

    for_response = ['移動源 C1', '固定燃燒源 C1', '工業製程 C1', '人為逸散 C1(工時計算)', '人為逸散 C1(冷媒設備)',
                    '人為逸散 C1(消防設備)', '其他關注類物質 C1', '輸入電力 C2(一般用電)',
                    '客戶和訪客運輸 C3', '員工差旅 C3', '員工通勤 C3', '營運產生之廢棄物 C4',
                    '資本設備 C4', '輸入能源上游排放 C4(輸入電力上游)', '用水(水資源管理) B.5.2.a', '購買商品 C4',
                    '上游租賃資產 C4', '顧問諮詢、清潔、維護等 C4', '上游運輸及配送 C3','廢棄物運輸 C3',
                    '下游運輸及配送 C3', '下游租賃資產 C5', '產品壽命終止階段 C5', '產品使用階段 C5',
                    '投資 C5', '其它間接排放 C6']

    file_path = resource_path('正確資料')
    url = url + "/project/calculate/"
    driver.get(url)
    index = 0
    for source, file_name in source_to_file_map.items():

        if source not in ['人為逸散 C1', '輸入電力 C2', '輸入蒸汽 C2']:
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, source)))
            link.click()

            excel_import(driver, source_to_file_map, file_path, file_name, url, index, for_response)
            index += 1
        elif source == '人為逸散 C1':
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, source)))
            link.click()
            time.sleep(1)
            excel_import(driver, source_to_file_map, file_path, file_name, url, index, for_response)
            index += 1

            driver.get(url + 'directFugitiveEmission')
            element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//div[text()="冷媒設備 B.2.2.d"]')))
            driver.execute_script("arguments[0].click()", element)
            time.sleep(2)
            excel_import(driver, source_to_file_map, file_path, '(新增)正確數值_RFG.xlsx', url, index, for_response)
            index += 1

            driver.get(url + 'directFugitiveEmission')
            time.sleep(1)
            element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//div[text()="消防設備 B.2.2.d"]')))
            driver.execute_script("arguments[0].click()", element)
            excel_import(driver, source_to_file_map, file_path, '(新增)正確數值_Fire.xlsx', url, index, for_response)
            index +=1
            time.sleep(2)
        elif source == '輸入電力 C2':
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, source)))
            link.click()
            time.sleep(1)
            page_height = driver.execute_script(
                "return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );")
            # 往下滾動至頁面中間
            driver.execute_script(f"window.scrollTo(0, {page_height // 9});")
            time.sleep(1)
            excel_import(driver, source_to_file_map, file_path, file_name, url, index, for_response)
            index +=1

        elif source == '輸入蒸汽 C2':
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, source)))
            link.click()
            excel_import(driver, source_to_file_map, file_path, file_name, url, index, for_response)
            index += 1
        elif source == '輸入能源上游排放 C4':
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, source)))
            link.click()
            time.sleep(2)
            excel_import(driver, source_to_file_map, file_path, file_name, url, index, for_response)
            index += 1

            driver.get(url + 'upstreamEmissions')
            time.sleep(1)
            element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//div[text()="其他輸入能源上游 B.5.2.a"]')))
            driver.execute_script("arguments[0].click()", element)
            excel_import(driver, source_to_file_map, file_path, '(新增)正確數值_輸入能源上游.xlsx', url, index, for_response)
            index += 1
            time.sleep(2)

    driver.get(url)

def excel_import(driver, source_to_file_map, file_path, file_name, url,index, for_response):
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '匯入 & 匯出')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(1)
    import_element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//ul//button[span[text()="匯入"]]')))
    driver.execute_script("arguments[0].click()", import_element)
    time.sleep(1)

    file_input = driver.find_element(By.XPATH, '//input[@id="Excel"]')  # 上傳檔案(路徑要自己改)
    temp = os.path.join(file_path, file_name)
    file_input.send_keys(temp)

    time.sleep(1)
    upload_element = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "button[class='ant-btn ant-btn-primary']")))
    driver.execute_script("arguments[0].click()", upload_element)  # 篩選上傳檔案有儲存按鈕的
    files_without_save_button = {
        '(新增)正確數值_Vehicle.xlsx', '(新增)正確數值_StationaryCombustionSource.xlsx', '(新增)正確數值_IndustrialProcess.xlsx',
        '(新增)正確數值_workhour.xlsx', '(新增)正確數值_other.xlsx', '(新增)正確數值_Elec.xlsx', 'GreenElec.xlsx', 'steam_plus.xlsx'
    }
    try:
        if file_name not in files_without_save_button:
            print(file_name)
            time.sleep(1)
            save_element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//span[contains(text(), "確 定")]')))
            driver.execute_script("arguments[0].click()", save_element)

            time.sleep(1)
            catch_response(driver, for_response[index])
            time.sleep(1)


        else:  # upload_path[now]=='RFG.xlsx' or upload_path[now]=='Fire.xlsx':
            print(file_name)
            # time.sleep(3)
            save_element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//button[@class="ant-btn css-dts6b9 ant-btn-primary"]')))
            driver.execute_script("arguments[0].click()", save_element)
            time.sleep(1)
            catch_response(driver, for_response[index])
    except:
        print("失敗")
    time.sleep(1)