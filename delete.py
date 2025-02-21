import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from response import catch_response


def delete(driver, url):
    source_type = ['移動源 C1', '固定燃燒源 C1', '工業製程 C1', '人為逸散 C1', '其他關注類物質 C1', '輸入電力 C2','輸入蒸汽 C2',
                   '客戶和訪客運輸 C3', '員工差旅 C3', '員工通勤 C3', '營運產生之廢棄物 C4', '資本設備 C4',
                   '輸入能源上游排放 C4', '用水(水資源管理) B.5.2.a', '購買商品 C4', '上游租賃資產 C4',
                   '顧問諮詢、清潔、維護等 C4', '上游運輸及配送 C3', '廢棄物運輸 C3', '下游運輸及配送 C3',
                   '下游租賃資產 C5', '產品壽命終止階段 C5', '產品使用階段 C5', '投資 C5', '其它間接排放 C6']

    filtered_elements = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.XPATH, '//div[@class="pficons"]/a')))
    text_values = []

    for element in filtered_elements:
        if element.get_attribute("disabled") != None:
            continue
        else:
            element_text = element.text
            text_values.append(element_text)

    url = url + "/calculate/"
    for i in source_type:
        if i == '人為逸散 C1':
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))
            link.click()
            time.sleep(1)
            delete_workhour(driver, url, i)

            element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//div[text()="冷媒設備 B.2.2.d"]')))
            driver.execute_script("arguments[0].click()", element)
            time.sleep(1)
            delete_data(driver, url, i, 1)

            element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//div[text()="消防設備 B.2.2.d"]')))
            driver.execute_script("arguments[0].click()", element)
            time.sleep(1)
            delete_data(driver, url, i, 2)
        elif i == '輸入能源上游排放 C4':
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))
            link.click()
            time.sleep(1)
            delete_data(driver, url, i, 3)

            element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//div[text()="其他輸入能源上游 B.5.2.a"]')))
            driver.execute_script("arguments[0].click()", element)
            time.sleep(1)
            delete_data(driver, url, i, 4)
        elif i == '輸入電力 C2':
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))
            link.click()
            delete_elec(driver, i, url)
            time.sleep(1)

        elif i == '輸入蒸汽 C2':
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))
            link.click()
            delete_steam(driver, url)
            time.sleep(1)
        else:
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))
            link.click()
            delete_data(driver, url, i, 0)


def delete_data(driver, url, i, rfg_or_fire):
    # 判斷是否還有需要刪除的
    while True:
        try:
            checkbox = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'span.anticon.anticon-delete')))
            driver.execute_script("arguments[0].click();", checkbox)
            time.sleep(1)
            checkbox = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, 'button.ant-btn.css-zjzpde.ant-btn-default.ant-btn-dangerous')))
            driver.execute_script("arguments[0].click();", checkbox)
            if rfg_or_fire == 1:
                catch_response(driver, '人為逸散 C1(冷媒設備)')
            elif rfg_or_fire == 2:
                catch_response(driver, '人為逸散 C1(消防設備)')
            elif rfg_or_fire == 3:
                catch_response(driver, '輸入電力上游 B.5.2.a')
            elif rfg_or_fire == 4:
                catch_response(driver, '其他輸入能源上游 B.5.2.a')
            elif rfg_or_fire == 0:
                catch_response(driver, i)
        except:
            if rfg_or_fire == 1:
                print("人為逸散 C1(冷媒設備) 沒有資料刪除!")
            elif rfg_or_fire == 2:
                print("人為逸散 C1(消防設備) 沒有資料刪除!")
            elif rfg_or_fire == 3:
                print("輸入電力上游 B.5.2.a 沒有資料刪除!")
            elif rfg_or_fire == 4:
                print("其他輸入能源上游 B.5.2.a 沒有資料刪除!")
            else:
                print(i + " 沒有資料刪除!")
            break


def delete_workhour(driver, i, url):
    while True:

        try:
            # 查找並點擊選擇框
            element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-form-item-control-input-content"]')))
            element.click()

            # 等待刪除圖標出現並點擊
            delete_icon = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//span[@class="anticon anticon-delete"]')))
            driver.execute_script("arguments[0].click();", delete_icon)

            # 等待確認刪除按鈕出現並點擊
            confirm_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, 'button.ant-btn.css-zjzpde.ant-btn-default.ant-btn-dangerous')))
            driver.execute_script("arguments[0].click();", confirm_button)

            # 捕獲刪除操作後的響應
            catch_response(driver, '工時計算 C1')  # 確保這個函數定義正確
        except:
            print("工時計算 C1 沒有資料刪除!")
            break


def delete_elec(driver, i, url):
    time.sleep(1)
    while True:
        try:
            # 確保每次操作前滾動頁面到中間
            page_height = driver.execute_script(
                "return Math.max(document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight);")
            driver.execute_script(f"window.scrollTo(0, {page_height // 3});")
            time.sleep(1)  # 確保滾動完成

            element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '(//div[@class="ant-form-item-control-input-content"])[1]')))
            element.click()
            delete_icon = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//span[@class="anticon anticon-delete"]')))
            driver.execute_script("arguments[0].click();", delete_icon)

            confirm_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, 'button.ant-btn.css-zjzpde.ant-btn-default.ant-btn-dangerous')))
            driver.execute_script("arguments[0].click();", confirm_button)
            catch_response(driver, '輸入電力 C1(一般用電)')
        except:
            driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(1)  # 確保滾動完成
            print("輸入電力 C1(一般用電) 沒有資料刪除!")
            break


def delete_steam(driver, url):
    bye_steam(driver, url, 0)

    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//div[text()="蒸氣減項 B.3.2.b"]')))
    driver.execute_script("arguments[0].click()", element)

    bye_steam(driver, url, 1)


def bye_steam(driver, url, plus_or_minus):
    while True:
        try:
            element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-form-item-control-input-content"]')))
            element.click()
            delete_icon = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//span[@class="anticon anticon-delete"]')))
            driver.execute_script("arguments[0].click();", delete_icon)

            confirm_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, 'button.ant-btn.css-zjzpde.ant-btn-default.ant-btn-dangerous')))
            driver.execute_script("arguments[0].click();", confirm_button)
            if plus_or_minus == 0:
                catch_response(driver, '輸入蒸汽 C1(蒸氣加項)')
            else:
                catch_response(driver, '輸入蒸汽 C1(蒸氣減項)')
        except:
            if plus_or_minus == 0:
                print("輸入蒸汽 C1(蒸氣加項) 沒有資料刪除")
            else:
                print("輸入蒸汽 C1(蒸氣減項) 沒有資料刪除")
            break