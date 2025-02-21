import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import logging
# 新增時輸入能源上游排放驗證(驗證評估項目與使用量)
def upstreamEmissions_add(driver,url, site_num):
    sign(driver, url,site_num)
    url = url+"/project/calculate/upstreamEmissions"
    driver.get(url)
    time.sleep(3)
    text = Add_verify(driver)
    specific_text = {'輸入電力上游' : 2400}
    for item, expected_quantity in specific_text.items():
        if item in text:
            actual_quantity = text[item]
            if actual_quantity == expected_quantity:
                print(f"評估項目 '{item}'，使用量: {actual_quantity} 正確")
                logging.info(f"評估項目 '{item}'，使用量: {actual_quantity} 正確")
            else:
                print(f"評估項目 '{item}'，使用量: {actual_quantity} 不正確 (正確使用量: {expected_quantity})")
                logging.info(f"評估項目 '{item}'，使用量: {actual_quantity} 不正確 (正確使用量: {expected_quantity})")
        else:
            print(f"評估項目 '{item}' 找不到")
            logging.info(f"評估項目 '{item}' 找不到")

    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="其他輸入能源上游 B.5.2.a"]')))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(3)
    text = Add_verify(driver)
    specific_text = {'輸入能源上游(ABC-123)' : 100,
                     '輸入能源上游(ABC-234)' : 100,
                     '輸入能源上游(ABC-345)' : 100,
                     '輸入能源上游(固定設備名稱測試1)' : 100,
                     '輸入能源上游(固定設備名稱測試2)' : 100,
                     '輸入能源上游(固定設備名稱測試3)' : 100,
                     '輸入能源上游(製程設備名稱測試1)' : 100,
                     '輸入能源上游(製程設備名稱測試2)' : 100,
                     '輸入能源上游(製程設備名稱測試3)' : 100}
    for item, expected_quantity in specific_text.items():
        if item in text:
            actual_quantity = text[item]
            if actual_quantity == expected_quantity:
                print(f"評估項目 '{item}'，使用量: {actual_quantity} 正確")
                logging.info(f"評估項目 '{item}'，使用量: {actual_quantity} 正確")
            else:
                print(f"評估項目 '{item}'，使用量: {actual_quantity} 不正確 (正確使用量: {expected_quantity})")
                logging.info(f"評估項目 '{item}'，使用量: {actual_quantity} 不正確 (正確使用量: {expected_quantity})")
        else:
            print(f"評估項目 '{item}' 找不到")
            logging.info(f"評估項目 '{item}' 找不到")
# 編輯後輸入能源上游排放驗證(評估項目,使用量,係數編號)
def upstreamEmissions_edit(driver,url, site_num):
    sign(driver, url, site_num)
    url = url+"/project/calculate/upstreamEmissions"

    driver.get(url)
    time.sleep(3)
    text = Edit_verify(driver)
    expected_values = {
        '輸入電力上游': {'quantity': 4800, 'coefficient': '【ER6002】'}
    }
    for item, expected  in expected_values.items():
        if item in text:
            entries = text[item]
            if len(entries) > 1:
                print(f"錯誤: 評估項目 '{item}' 出現了多次 ({len(text[item])} 次)，僅允許出現一次。")
                logging.info(f"錯誤: 評估項目 '{item}' 出現了多次 ({len(text[item])} 次)，僅允許出現一次。")
            else:
                actual = entries[0]
                if (actual['quantity'] == expected['quantity'] and
                    actual['coefficient'] == expected['coefficient']):
                    print(f"評估項目 '{item}',使用量: {actual['quantity']} 和 係數編號: {actual['coefficient']} 正確")
                    logging.info(f"評估項目 '{item}',使用量: {actual['quantity']} 和 係數編號: {actual['coefficient']} 正確")
                else:
                    print(f"評估項目 '{item}' 的結果不正確")
                    logging.info(f"評估項目 '{item}' 的結果不正確")
                    print(f"實際使用量: {actual['quantity']}，預期使用量: {expected['quantity']}")
                    logging.info(f"實際使用量: {actual['quantity']}，預期使用量: {expected['quantity']}")
                    print(f"實際係數編號: {actual['coefficient']}，預期係數編號: {expected['coefficient']}")
                    logging.info(f"實際係數編號: {actual['coefficient']}，預期係數編號: {expected['coefficient']}")
        else:
            print(f"評估項目 '{item}' 找不到")
            logging.info(f"評估項目 '{item}' 找不到")

    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="其他輸入能源上游 B.5.2.a"]')))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(3)
    text = Edit_verify(driver)
    specific_text =  {
         '輸入能源上游(ABC-234)': {'quantity': 100, 'coefficient': '-'},
         '輸入能源上游(ABC-345)': {'quantity': 100, 'coefficient': '-'},
         '輸入能源上游(ABC-789)': {'quantity': 50, 'coefficient': '【ER3002】'},
         '輸入能源上游(ABC-123)': {'quantity': 100, 'coefficient': '-'},
         '輸入能源上游(固定設備名稱測試編輯)': {'quantity': 50, 'coefficient': '-'},
         '輸入能源上游(固定設備名稱測試2)': {'quantity': 100, 'coefficient': '-'},
         '輸入能源上游(固定設備名稱測試3)': {'quantity': 100, 'coefficient': '-'},
         '輸入能源上游(製程設備名稱測試編輯)': {'quantity': 50, 'coefficient': '-'},
         '輸入能源上游(製程設備名稱測試2)': {'quantity': 100, 'coefficient': '-'},
         '輸入能源上游(製程設備名稱測試3)': {'quantity': 100, 'coefficient': '-'},
     }

    for item, expected_quantity in specific_text.items():
        if item in text:
            entries = text[item]
            if len(entries) > 1:
                print(f"錯誤: 評估項目 '{item}' 出現了多次 ({len(text[item])} 次)，僅允許出現一次。")
            else:
                actual = entries[0]
                if (actual['quantity'] == expected_quantity['quantity'] and
                    actual['coefficient'] == expected_quantity['coefficient']):
                    print(f"評估項目 '{item}',使用量: {actual['quantity']} 和 係數編號: {actual['coefficient']} 正確")
                    logging.info(f"評估項目 '{item}',使用量: {actual['quantity']} 和 係數編號: {actual['coefficient']} 正確")
                else:
                    print(f"評估項目 '{item}' 的結果不正確")
                    logging.info(f"評估項目 '{item}' 的結果不正確")
                    print(f"實際使用量: {actual['quantity']}，預期使用量: {expected_quantity['quantity']}")
                    logging.info(f"實際使用量: {actual['quantity']}，預期使用量: {expected_quantity['quantity']}")
                    print(f"實際係數編號: {actual['coefficient']}，預期係數編號: {expected_quantity['coefficient']}")
                    logging.info(f"實際係數編號: {actual['coefficient']}，預期係數編號: {expected_quantity['coefficient']}")
        else:
            print(f"評估項目 '{item}' 找不到")
            logging.info(f"評估項目 '{item}' 找不到")
#匯入後能源上游驗證
def upstreamEmissions_input(driver,url, site_num):
    sign(driver, url,site_num)
    url = url+"/project/calculate/upstreamEmissions"
    driver.get(url)
    time.sleep(3)
    text = Add_verify(driver)
    specific_text = {'輸入電力上游' : 2400}
    for item, expected_quantity in specific_text.items():
        if item in text:
            actual_quantity = text[item]
            if actual_quantity == expected_quantity:
                print(f"評估項目 '{item}'，使用量: {actual_quantity} 正確")
                logging.info(f"評估項目 '{item}'，使用量: {actual_quantity} 正確")
            else:
                print(f"評估項目 '{item}'，使用量: {actual_quantity} 不正確 (正確使用量: {expected_quantity})")
                logging.info(f"評估項目 '{item}'，使用量: {actual_quantity} 不正確 (正確使用量: {expected_quantity})")
        else:
            print(f"評估項目 '{item}' 找不到")
            logging.info(f"評估項目 '{item}' 找不到")

    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="其他輸入能源上游 B.5.2.a"]')))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(3)
    text = Add_verify(driver)
    specific_text = {'輸入能源上游(ABC-123)' : 100,
                     '輸入能源上游(ABC-234)' : 100,
                     '輸入能源上游(ABC-345)' : 100,
                     '輸入能源上游(固定設備名稱測試1)' : 100,
                     '輸入能源上游(固定設備名稱測試2)' : 100,
                     '輸入能源上游(固定設備名稱測試3)' : 100,
                     '輸入能源上游(製程設備名稱測試1)' : 100,
                     '輸入能源上游(製程設備名稱測試2)' : 100,
                     '輸入能源上游(製程設備名稱測試3)' : 100}
    for item, expected_quantity in specific_text.items():
        if item in text:
            actual_quantity = text[item]
            if actual_quantity == expected_quantity:
                print(f"評估項目 '{item}'，使用量: {actual_quantity} 正確")
                logging.info(f"評估項目 '{item}'，使用量: {actual_quantity} 正確")
            else:
                print(f"評估項目 '{item}'，使用量: {actual_quantity} 不正確 (正確使用量: {expected_quantity})")
                logging.info(f"評估項目 '{item}'，使用量: {actual_quantity} 不正確 (正確使用量: {expected_quantity})")
        else:
            print(f"評估項目 '{item}' 找不到")
            logging.info(f"評估項目 '{item}' 找不到")
# 新增後的資料抓取
def Add_verify(driver):
    carbon_emissions_dict = {}
    
    while True:
        try:
            # 等待表格加载完成
            table = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
            
            # 定位表头行元素
            thead_row = table.find_element(By.XPATH, './/thead/tr')
            
            # 找到 "評估項目" 和 "使用量" 在表头中的索引位置
            item_index = -1
            quantity_index = -1
            header_cells = thead_row.find_elements(By.TAG_NAME, 'th')
            for i, cell in enumerate(header_cells):
                if cell.text == '評估項目':
                    item_index = i
                elif cell.text == '使用量':
                    quantity_index = i
            
            # 检查是否找到 "評估項目" 和 "使用量" 的索引位置
            if item_index != -1 and quantity_index != -1:
                # 定位所有的 <tbody> 元素
                tbodies = table.find_elements(By.TAG_NAME, 'tbody')
                
                # 遍历每个 <tbody>
                for tbody in tbodies:
                    # 定位该 <tbody> 中所有的行元素
                    rows = tbody.find_elements(By.TAG_NAME, 'tr')
                    
                    # 遍历每个行元素
                    for row in rows[1:]:
                        # 定位“評估項目”和“使用量”元素
                        item_elements = row.find_elements(By.XPATH, f'.//td[{item_index + 1}]')
                        quantity_elements = row.find_elements(By.XPATH, f'.//td[{quantity_index + 1}]')
                        
                        # 获取“評估項目”和“使用量”的值
                        for item_element, quantity_element in zip(item_elements, quantity_elements):
                            item_text = item_element.text.strip()
                            quantity_text = quantity_element.text.strip()
                            if item_text and quantity_text:
                                try:
                                    quantity = float(quantity_text)  # 将使用量转为浮点数
                                    carbon_emissions_dict[item_text] = quantity
                                except ValueError:
                                    print(f"無法將使用量 '{quantity_text}' 轉換為浮點數")
            
            # 检查是否有下一页的按钮，如果有则点击并继续，否则结束循环
            try:
                next_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//a[text()="Next"]')))
                next_button.click()
                time.sleep(3)  # 等待页面加载
            except:
                break  # 没有下一页按钮，结束循环
        
        except Exception as e:
            print(f"處理表格時發生錯誤：{str(e)}")
            break
    
    return carbon_emissions_dict
# 編輯後的資料抓取
def Edit_verify(driver):
    carbon_emissions_dict = {}
    
    while True:
        try:
            # 等待表格加载完成
            table = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
            
            # 定位表头行元素
            thead_row = table.find_element(By.XPATH, './/thead/tr')
            
            # 找到 "評估項目" 和 "使用量" 在表头中的索引位置
            item_index = -1
            quantity_index = -1
            coefficient_index = -1
            header_cells = thead_row.find_elements(By.TAG_NAME, 'th')
            for i, cell in enumerate(header_cells):
                if cell.text == '評估項目':
                    item_index = i
                elif cell.text == '使用量':
                    quantity_index = i
                elif cell.text == '係數編號':
                    coefficient_index = i
                    
            
            # 检查是否找到 "評估項目" 和 "使用量" ,"係數編號"的索引位置
            if item_index != -1 and quantity_index != -1 and coefficient_index != -1:
                # 定位所有的 <tbody> 元素
                tbodies = table.find_elements(By.TAG_NAME, 'tbody')
                
                # 遍历每个 <tbody>
                for tbody in tbodies:
                    # 定位该 <tbody> 中所有的行元素
                    rows = tbody.find_elements(By.TAG_NAME, 'tr')
                    
                    # 遍历每个行元素
                    for row in rows[1:]:
                        # 定位“評估項目”和“使用量”元素
                        item_elements = row.find_elements(By.XPATH, f'.//td[{item_index + 1}]')
                        quantity_elements = row.find_elements(By.XPATH, f'.//td[{quantity_index + 1}]')
                        coefficient_elements =  row.find_elements(By.XPATH, f'.//td[{coefficient_index + 1}]')
                        
                        # 获取“評估項目”和“使用量”的值
                        for item_element, quantity_element,coefficient_element in zip(item_elements, quantity_elements, coefficient_elements):
                            item_text = item_element.text.strip()
                            quantity_text = quantity_element.text.strip()
                            coefficient_text = coefficient_element.text.strip()
                            
                            if item_text and quantity_text and coefficient_text:
                                try:
                                    quantity = float(quantity_text)  # 将使用量转为浮点数
                                    if item_text not in carbon_emissions_dict:
                                        carbon_emissions_dict[item_text] = []  # 创建一个空列表来存储同名项目
                                    carbon_emissions_dict[item_text].append({
                                                                            "quantity": quantity,
                                                                            "coefficient": coefficient_text
                                                                            })
                                except ValueError:
                                    print(f"無法將使用量 '{quantity_text}' 轉換為浮點數")

            break
            
            # 检查是否有下一页的按钮，如果有则点击并继续，否则结束循环
            try:
                next_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//a[text()="Next"]')))
                next_button.click()
                time.sleep(3)  # 等待页面加载
            except:
                break  # 没有下一页按钮，结束循环
        
        except Exception as e:
            print(f"處理表格時發生錯誤：{str(e)}")
            break
    return carbon_emissions_dict
# 導入排放源頁面
def sign(driver, url,site_num):
    site_dict = {1: "https://prod.netzero.com.tw", 2: "https://demo2.netzero.com.tw",
                 3: "https://14064-1.com", 4: "https://ghg.netzeroyun.com",
                 5: "https://220.132.206.5:666", 6: "https://220.132.206.5:8888",
                 7: "https://220.132.206.5:9999"}
    url = site_dict[site_num]
    time.sleep(2)
    driver.get(url)
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

    button = parent_tr.find_element(By.XPATH,'.//button[@type="button" and contains(@class, "ant-btn ant-btn-round ant-btn-link ant-btn-icon-only text-primary  border-0")]')
    # JavaScript寫法
    driver.execute_script("arguments[0].click()", button)
    time.sleep(1)

    link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "排放資料輸入")))  # 點擊排放資料輸入
    link.click()
                            