import time
from calendar import month
from enum import verify
from numpy.testing.print_coercion_tables import print_coercion_table
from selenium.webdriver.common.by import By
from selenium.webdriver.common.devtools.v128.page import print_to_pdf
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from decimal import Decimal,ROUND_HALF_UP

from unicodedata import decimal


def value_test_45(driver, url, site_num):
    source_type = ['移動源 C1', '固定燃燒源 C1', '工業製程 C1', '人為逸散 C1','其他關注類物質 C1', '輸入電力 C2', '客戶和訪客運輸 C3',
                   '員工差旅 C3', '員工通勤 C3', '營運產生之廢棄物 C4', '資本設備 C4', '輸入能源上游排放 C4',
                   '用水(水資源管理) B.5.2.a', '購買商品 C4', '上游租賃資產 C4', '顧問諮詢、清潔、維護等 C4',
                   '上游運輸及配送 C3', '廢棄物運輸 C3', '下游運輸及配送 C3', '下游租賃資產 C5', '產品壽命終止階段 C5',
                   '產品使用階段 C5', '投資 C5', '其它間接排放 C6']
    test_source = ['mobileCombustion', 'stationaryCombustion', 'directProcessEmission', 'directFugitiveEmission', 'otherCompound',
                   'electricity', 'visitor', 'businessTravel', 'commuting', 'disposal',
                   'capitalGood', 'upstreamEmissions', 'waterUsage', 'purchasedGood', 'leasedAsset', 'consultant',
                   'upstreamTransport', 'disposalDownTransport',
                   'downstreamTransport', 'downLeasedAsset', 'downstreamDisposal', 'useEmission', 'investment', 'other']

    url = url + "/project/calculate/"
    for i, j in zip(source_type, test_source):
        if j in {'mobileCombustion', 'stationaryCombustion', 'directProcessEmission'}:
            navigate_to_page(driver, url + j)
            calculation_process_button(driver,i)  # 驗算"計算過程"
            print("[[[[       每筆加總與總值驗證       ]]]]")
            print(i + " 每筆相加的 CO2排放量")
            total_co2 = calculate_co2_mobileCombustion(driver)  # 每筆碳排放量加總
            print(i + " 總值的 CO2排放量")
            co2_validate = validate_co2_mobileCombustion(driver)  # 總計kgco2e
            if f"{total_co2:.4f}" == f"{co2_validate:.4f}":
                print("[ CO2相等: 驗證正確!!! ✅✅]")
            else:
                print("[ CO2不相等: 驗證錯誤!!! ❌❌]")
        elif j =='directFugitiveEmission' :
            driver.get(url + j)
            all_co2 = 0
            element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
            element.click()
            time.sleep(2)
            options = WebDriverWait(driver,10).until(EC.presence_of_all_elements_located((By.XPATH,'//div[contains(@class,"ant-select-item-option-content")]')))
            time.sleep(2)
            for index in range(len(options)):
                options = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located(
                    (By.XPATH, '//div[contains(@class,"ant-select-item-option-content")]')))
                time.sleep(2)
                options[index].click()
                calculation_process_button_workhour(driver, i)
                total_co2 = calculate_co2_workhour(driver)
                all_co2 += total_co2
                time.sleep(2)
                driver.get(url + j)
                time.sleep(3)
                element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH,
                                                                                      '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
                element.click()

            all_co2 = round(all_co2,4)
            print("[[[[       每筆加總與總值驗證       ]]]]")
            print("工時計算 B.2.2.d 每筆相加的 CO2排放量")
            print(f'碳排放量總和：{all_co2}')  # 計算Co2
            print("工時計算 B.2.2.d 總值的 CO2排放量")
            co2_validate = validate_co2_visitor(driver)  # 下方總值計算
            # 表單內計算與下方總值驗證
            if f"{all_co2:.4f}" == f"{co2_validate:.4f}":
                print("[ CO2相等: 驗證正確!!! ✅✅ ]")
            else:
                print("[ CO2不相等: 驗證錯誤!!! ❌❌ ]")
            print('\n')
            time.sleep(2)

            driver.get(url+j)
            time.sleep(5)
            element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//div[text()="冷媒設備 B.2.2.d"]')))
            driver.execute_script("arguments[0].click()", element)
            page_height = driver.execute_script(
                "return Math.max(document.body.scrollHeight, document.documentElement.scrollHeight);"
            )
            driver.execute_script(f"window.scrollTo(0, {page_height // 3});")
            time.sleep(3)
            calculation_process_button_rfg(driver, i)  # 驗算"計算過程"
            print("[[[[       每筆加總與總值驗證       ]]]]")
            print("冷媒設備 B.2.2.d 每筆相加的 CO2排放量")
            total_co2 = calculate_co2_visitor(driver)  # 每筆碳排放量加總
            print("冷媒設備 B.2.2.d 總值的 CO2排放量")
            co2_validate = validate_co2_visitor(driver)  # 總計kgco2e
            if f"{total_co2:.4f}" == f"{co2_validate:.4f}":
                print("[ CO2相等: 驗證正確!!! ✅✅]")
            else:
                print("[ CO2不相等: 驗證錯誤!!! ❌❌]")
            print('\n')
            time.sleep(2)

            driver.get(url+j)
            time.sleep(5)
            element = WebDriverWait(driver,10).until(
                EC.element_to_be_clickable((By.XPATH,'//div[text()="消防設備 B.2.2.d"]'))
            )
            driver.execute_script("arguments[0].click()", element)
            time.sleep(2)
            page_height = driver.execute_script(
                "return Math.max(document.body.scrollHeight, document.documentElement.scrollHeight);"
            )
            driver.execute_script(f"window.scrollTo(0, {page_height // 3});")
            time.sleep(2)
            calculation_process_button_fir(driver, i)  # 驗算"計算過程"
            print("[[[[       每筆加總與總值驗證       ]]]]")
            print("消防設備 B.2.2.d 每筆相加的 CO2排放量")
            total_co2 = calculate_co2_visitor(driver)  # 每筆碳排放量加總
            print("消防設備 B.2.2.d 總值的 CO2排放量")
            co2_validate = validate_co2_visitor(driver)  # 總計kgco2e
            if f"{total_co2:.4f}" == f"{co2_validate:.4f}":
                print("[ CO2相等: 驗證正確!!!✅✅ ]")
            else:
                print("[ CO2不相等: 驗證錯誤!!!❌❌ ]")
            print('\n')
            time.sleep(2)
        elif j == 'otherCompound':
            navigate_to_page(driver, url+j)
            calculation_process_button_other(driver, i)  # 驗算"計算過程"
            print("[[[[       每筆加總與總值驗證       ]]]]")
            print("其他關注類物質 C1 每筆相加的 CO2排放量")
            total_co2 = calculate_co2_visitor(driver)  # 每筆碳排放量加總
            print("其他關注類物質 C1 總值的 CO2排放量")
            co2_validate = validate_co2_visitor(driver)  # 總計kgco2e
            if f"{total_co2:.4f}" == f"{co2_validate:.4f}":
                print("[ CO2相等: 驗證正確!!!✅✅ ]")
            else:
                print("[ CO2不相等: 驗證錯誤!!!❌❌ ]")
            print('\n')
            time.sleep(2)
        elif j == 'electricity':
            driver.get(url+j)
            all_co2 = 0
            time.sleep(2)
            page_height = driver.execute_script(
                "return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );")
            # 往下滾動至頁面中間
            driver.execute_script(f"window.scrollTo(0, {page_height // 3});")
            time.sleep(2)
            element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH,
                                                                                  '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
            element.click()
            time.sleep(2)
            options = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located(
                (By.XPATH, '//div[contains(@class,"ant-select-item-option-content")]')))
            time.sleep(2)
            for index in range(len(options)):
                options = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located(
                    (By.XPATH, '//div[contains(@class,"ant-select-item-option-content")]')))
                time.sleep(2)
                options[index].click()
                calculation_process_button_elec(driver, i)
                total_co2 = calculate_co2_workhour(driver)
                all_co2 += total_co2
                time.sleep(2)
                driver.get(url+j)
                time.sleep(2)
                page_height = driver.execute_script(
                    "return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );")
                # 往下滾動至頁面中間
                driver.execute_script(f"window.scrollTo(0, {page_height // 3});")
                time.sleep(2)
                element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH,
                                                                                      '//div[@class="ant-select ant-select-in-form-item ant-select-single ant-select-show-arrow"]')))
                element.click()
            all_co2 = round(all_co2,4)
            print("[[[[       每筆加總與總值驗證       ]]]]")
            print("一般用電 B.3.2.a 每筆相加的 CO2排放量")
            print(f'碳排放量總和：{all_co2}')  # 計算Co2
            print("一般用電 B.3.2.a 總值的 CO2排放量")
            co2_validate = validate_co2_workhour(driver)
            if f"{all_co2:.2f}" == f"{co2_validate:.2f}":
                print("[ CO2相等: 驗證正確!!! ✅✅ ]")
            else:
                print("[ CO2不相等: 驗證錯誤!!! ❌❌ ]")
            print('\n')
            time.sleep(2)
        elif j == 'upstreamEmissions' :
            navigate_to_page(driver, url + j)
            calculation_process_button_c3(driver, i,1)  # 驗算"計算過程"
            print("[[[[       每筆加總與總值驗證       ]]]]")
            print("輸入電力上游 B.5.2.a 每筆相加的 CO2排放量")
            total_co2 = calculate_co2_visitor(driver)  # 每筆碳排放量加總
            print("輸入電力上游 B.5.2.a 總值的驗證 CO2排放量")
            co2_validate = validate_co2_visitor_c3(driver)  # 總計kgco2e
            if f"{total_co2:.4f}" == f"{co2_validate:.4f}":
                print("[ CO2相等: 驗證正確!!! ✅✅ ]")
            else:
                print("[ CO2不相等: 驗證錯誤!!! ❌❌ ]")
            print('\n')
            time.sleep(2)

            driver.get(url+j)
            time.sleep(5)
            element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//div[text()="其他輸入能源上游 B.5.2.a"]')))
            driver.execute_script("arguments[0].click()", element)
            page_height = driver.execute_script(
                "return Math.max(document.body.scrollHeight, document.documentElement.scrollHeight);"
            )
            driver.execute_script(f"window.scrollTo(0, {page_height // 3});")
            time.sleep(2)
            calculation_process_button_c3(driver, i,2)  # 驗算"計算過程"
            print("[[[[       每筆加總與總值驗證       ]]]]")
            print("其他輸入能源上游 B.5.2.a 每筆相加的 CO2排放量")
            total_co2 = calculate_co2_visitor(driver)  # 每筆碳排放量加總
            print("輸入電力上游 B.5.2.a 總值的驗證 CO2排放量")
            co2_validate = validate_co2_visitor_c3(driver)  # 總計kgco2e
            if f"{total_co2:.4f}" == f"{co2_validate:.4f}":
                print("[ CO2相等: 驗證正確!!! ✅✅ ]")
            else:
                print("[ CO2不相等: 驗證錯誤!!! ❌❌ ]")
            print('\n')
            time.sleep(2)
        else:
            navigate_to_page(driver, url + j)
            calculation_process_button_c3(driver, i,0)  # 驗算"計算過程"
            print("[[[[       每筆加總與總值驗證       ]]]]")
            print(i + " 每筆相加的 CO2排放量")
            total_co2 = calculate_co2_visitor(driver)  # 每筆碳排放量加總
            print(i + " 總值的驗證 CO2排放量")
            co2_validate = validate_co2_visitor_c3(driver)  # 總計kgco2e
            if f"{total_co2:.4f}" == f"{co2_validate:.4f}":
                print("[ CO2相等: 驗證正確!!! ✅✅ ]")
            else:
                print("[ CO2不相等: 驗證錯誤!!! ❌❌ ]")
            print('\n')
            time.sleep(2)


# 進入指定排放源,並滾動至底
def navigate_to_page(driver, url):
    driver.get(url)
    time.sleep(5)
    page_height = driver.execute_script(
        "return Math.max(document.body.scrollHeight, document.documentElement.scrollHeight);"
    )
    driver.execute_script(f"window.scrollTo(0, {page_height // 3});")
    time.sleep(2)

# 點擊計算過程按鈕(移動，固定，製程)
def calculation_process_button(driver, i):
    has_next_page = True
    while has_next_page:
        # 定位要滾動的table
        table_scroll_container = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//div[contains(@class, "ant-table-content")]'))
        )

        # 向右滾動到最底
        driver.execute_script("arguments[0].scrollLeft = arguments[0].scrollWidth;", table_scroll_container)
        time.sleep(2)

        # 定位table表單
        table = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
        thead_row = table.find_element(By.XPATH, './/thead/tr')

        # 對應 "碳排放量(kgCO2e)"
        carbon_emissions_index = -1
        header_cells = thead_row.find_elements(By.TAG_NAME, 'th')
        for idx, cell in enumerate(header_cells):
            if cell.text == '碳排放量(kgCO2e)':
                carbon_emissions_index = idx
                break

        # 找到所有包含"計算過程"的按鈕
        buttons = table.find_elements(By.XPATH, './/td[@class="ant-table-cell calculate_table"]//span[text()="計算過程"]')
        if not buttons:
            print("找不到'計算過程'按鈕")
            return
        # 存取每筆碳排放量
        carbon_emissions_dict = {}
        print(f"{i} 計算結果 :")
        for index_2, button in enumerate(buttons):
            print(f"{i} - 第 {index_2 + 1} 筆 驗證:")
            # 找到與當前按鈕對應的"碳排放量(kgCO2e)"
            row = button.find_element(By.XPATH, "./ancestor::tr")  # 找到該按鈕所在的 `tr`
            carbon_emissions_element = row.find_element(By.XPATH, f'.//td[{carbon_emissions_index + 1}]')  # 定位該筆的碳排放量
            try:
                carbon_emissions = float(carbon_emissions_element.text)
                carbon_emissions_dict[index_2 + 1] = carbon_emissions
            except :
                print("無資料")
                continue
            # 點擊按鈕
            driver.execute_script("arguments[0].scrollIntoView();", button)  # 確保按鈕可見
            driver.execute_script("arguments[0].click();", button)  # 以 JavaScript 點擊，避免元素被遮擋
            time.sleep(2)

            # 執行計算結果驗證
            process_table_data_and_verify(driver, i,index_2,carbon_emissions_dict)
        try:
            next_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, "//li[@title='下一頁']//button[@class='ant-pagination-item-link']")))
            ActionChains(driver).move_to_element(next_button).click().perform()
            time.sleep(2)
        except Exception as e:
            has_next_page = False

    return carbon_emissions_dict

# 抓取排放係數、使用量、使用量*排放係數,並計算結果與驗證(移動,固定,製程)
def process_table_data_and_verify(driver,i,index_2,carbon_emissions_dict):
    row_names = {0: "CO2", 1: "CH4", 2: "N2O"}
    results = {"CO2": [], "CH4": [], "N2O": []}
    usage = None
    GWP = {}
    KgCO2e = {}
    total_KgCO2e = {}
    calculated_results = {}
    verification_results = {}  # 儲存驗證結果

    table = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f'(//div[contains(@class, "ant-table-content")])[{index_2 + 2}]')))
    header_index = get_column_index(table, "排放係數")
    # header_index應該為3
    if header_index == -1:
        print("未找到 '排放係數' 欄位")
        return
    # 定位table中的tbody/tr元素
    rows = table.find_elements(By.XPATH, './/tbody/tr[@class="ant-table-row ant-table-row-level-0"]')
    for index, row in enumerate(rows):
        row_name = row_names.get(index, f"Row_{index}")
        if index == 0:
            value_td_index = header_index + 1 # CO2排放係數
            usage_td_index = header_index + 2 # 使用量
            verify_td_index = header_index + 3  # 使用量*排放係數
            GWP_td_index = header_index + 4 # GWP
            KgCO2e_td_index = header_index + 5 # 排放量KgCO2e
            total_KgCO2e_td_index = header_index + 6 # 總計KgCO2e
            # 定位各欄位的值
            usage_elements = row.find_elements(By.XPATH, f'.//td[{usage_td_index}]')
            verify_elements = row.find_elements(By.XPATH, f'.//td[{verify_td_index}]')
            GWP_elements = row.find_elements(By.XPATH, f'.//td[{GWP_td_index}]')
            kgco2e_elements = row.find_elements(By.XPATH,f'.//td[{KgCO2e_td_index}]')
            total_kgco2e_elements = row.find_elements(By.XPATH, f'.//td[{total_KgCO2e_td_index}]')

            # 使用量取值
            if usage_elements:
                usage = float(usage_elements[0].text)
            # 使用量*排放係數取值
            if verify_elements:
                verify_value = float(verify_elements[0].text)
            # 取得"CO2 使用量*排放係數",並存入verification_results字典
                verification_results[row_name] = verify_value
            # GWP取值
            if GWP_elements:
                GWP_valus = float(GWP_elements[0].text)
                # 將取得的GWP存入 GWP_valus字典裡
                GWP[row_name] = GWP_valus
            # kgco2e取值
            if kgco2e_elements:
                kgco2e_valus = float(kgco2e_elements[0].text)
                KgCO2e[row_name] = kgco2e_valus
            # 總計kgco2e取值
            if total_kgco2e_elements:
                total_KgCO2e_valus = float(total_kgco2e_elements[0].text)
                total_KgCO2e[row_name] = total_KgCO2e_valus
        # CH4與N2O取值
        elif index == 1 or index == 2:
            value_td_index = header_index # CH4與N2O 排放係數
            verify_td_index = header_index +1 # CH4與N2O 使用量x排放係數
            GWP_td_index = header_index + 2 #CH4與N2O GWP
            KgCO2e_td_index = header_index + 3 # CH4與N2O 排放量kgco2e

            verify_elements = row.find_elements(By.XPATH, f'.//td[{verify_td_index}]')
            GWP_elements = row.find_elements(By.XPATH, f'.//td[{GWP_td_index}]')
            kgco2e_elements = row.find_elements(By.XPATH, f'.//td[{KgCO2e_td_index}]')
            if verify_elements:
                verify_value = float(verify_elements[0].text)
                #取得 "N2O與CH4的使用量*排放係數",並存入verification_results字典
                verification_results[row_name] = verify_value
            if GWP_elements:
                GWP_valus = float(GWP_elements[0].text)
                GWP[row_name] = GWP_valus
            if kgco2e_elements:
                kgco2e_valus = float(kgco2e_elements[0].text)
                KgCO2e[row_name] = kgco2e_valus
        else:
            continue
        td_elements = row.find_elements(By.XPATH, f'.//td[{value_td_index}]')

        for td in td_elements:
            results[row_name].append(float(td.text))
    # 計算排放係數*使用量
    if usage:
        for gas, emissions in results.items():
            calculated_results[gas] = round(sum(emissions) * usage, 4)

    # 驗證計算結果是否與頁面一致
    verify_calculated_results(calculated_results, verification_results,GWP,KgCO2e,total_KgCO2e)
    try:
        first_total_kgco2e = list(total_KgCO2e.values())[0]  # 取得 total_KgCO2e 第一個值
        first_carbon_emission = carbon_emissions_dict.get(index_2 + 1)  # 取得 carbon_emissions_dict 第 index_2+1 筆
        print(f"計算結果內的總計: {first_total_kgco2e} \n單筆碳排放量 : {first_carbon_emission}")

        if abs(first_total_kgco2e - first_carbon_emission) < 0.001:  # 允許微小誤差
            print("計算結果內的總計與單筆碳排放量 : 一致 ✅")
        else:
            print("計算結果內的總計與單筆碳排放量 : 不一致 ❌")

    except :
        print("無資料")

    time.sleep(2)
    # 當驗證完成後,點擊取消,關閉視窗
    try:
        modals = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, '//div[@class="ant-modal-content"]'))
        )
        # 取最後一個出現的 modal
        latest_modal = modals[-1]
        # 在這個 modal 內找到取消按鈕
        cancel_button = WebDriverWait(latest_modal, 10).until(
            EC.element_to_be_clickable((By.XPATH, './/button[@class="ant-btn ant-btn-default"]//span[text()="取 消"]'))
        )
        cancel_button.click()
        time.sleep(2)
    except Exception as e:
        print(f"無點擊到取消")

# 驗證計算結果是否與頁面一致
def verify_calculated_results(calculated_results, verification_results, GWP, KgCO2e,total_KgCO2e):
    total_calculated_kgco2e = 0
    for gas, calculated_value in calculated_results.items():
        verify_value = verification_results.get(gas)
        gwp_value = GWP.get(gas)
        kgco2e_value = KgCO2e.get(gas)

        # 驗證 使用量 * 排放係數
        if verify_value is not None:
            if abs(calculated_value - verify_value) < 1e-6:  # 相減時容許最小誤差值
                print(f"{gas} 使用量*排放係數 : 程式計算結果: {calculated_value}, 頁面呈現結果: {verify_value} 一致 ✅")
            else:
                print(f"{gas} 使用量*排放係數 : 程式計算結果: {calculated_value}, 頁面呈現結果: {verify_value} 不一致!!! ❌")
        else:
            print(f"無 {gas}")

        # 驗證 (使用量*排放係數) * GWP 是否等於 KgCO2e
        if verify_value is not None and gwp_value is not None and kgco2e_value is not None:
            expected_kgco2e = round(verify_value * gwp_value, 4)  # 程式計算 KgCO2e
            if abs(expected_kgco2e - kgco2e_value) < 1e-6:  # 相減時容許最小誤差值
                print(f"{gas} 排放量KgCO2e 驗證 : 程式計算結果: {expected_kgco2e}, 頁面呈現結果: {kgco2e_value} 一致 ✅")
            else:
                print(f"{gas} 排放量KgCO2e 驗證 : 程式計算結果: {expected_kgco2e}, 頁面呈現結果: {kgco2e_value} 不一致!!! ❌")

            total_calculated_kgco2e += kgco2e_value # 計算過程 - 排放量KgCO2e加總
        else:
            print(f"無 {gas}")

    # 驗證 排放量KgCO2e加總 是否等於 總計KgCO2e
    if total_KgCO2e:
        total_kgco2e_value = total_KgCO2e.get("CO2", 0)  # CO2 行可能存放總計數值(頁面抓取的值)
        if total_kgco2e_value is not None:
            total_calculated_kgco2e = round(total_calculated_kgco2e, 4)  # 確保為4位數
            if abs(total_calculated_kgco2e - total_kgco2e_value) < 1e-6:   # 相減時容許最小誤差值
                print(f"總計KgCO2e 驗證 : 程式計算結果: {total_calculated_kgco2e}, 頁面呈現結果: {total_kgco2e_value} 一致 ✅")
            else:
                print(
                    f"總計KgCO2e 驗證 : 程式計算結果: {total_calculated_kgco2e}, 頁面呈現結果: {total_kgco2e_value} 不一致!!! ❌")
        else:
            print("無 總計KgCO2e")

# 定位表頭欄位
def get_column_index(table, column_name):
    thead_row = table.find_element(By.XPATH, './/thead/tr')
    header_cells = thead_row.find_elements(By.TAG_NAME, 'th')
    for i, cell in enumerate(header_cells):
        if cell.text == column_name:
            return i
    return -1

# 判斷是否有下一頁,若有則點擊
def click_next_page(driver):
    try:
        next_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
            (By.XPATH, "//li[@title='下一頁']//button[@class='ant-pagination-item-link']")))
        ActionChains(driver).move_to_element(next_button).click().perform()
        time.sleep(2)
        return True
    except Exception:
        return False


def calculate_co2_mobileCombustion(driver):
    total_carbon_emissions = Decimal('0.0000')
    has_next_page = True
    while has_next_page:
        table = driver.find_element(By.TAG_NAME, 'table')
        thead_row = table.find_element(By.XPATH, './/thead/tr')

        # 找到 "碳排放量(kgCO2e)" 在表頭中的索引位置
        carbon_emissions_index = -1
        header_cells = thead_row.find_elements(By.TAG_NAME, 'th')
        for i, cell in enumerate(header_cells):
            if cell.text == '碳排放量(kgCO2e)':
                carbon_emissions_index = i

        # 檢查是否找到 "碳排放量(kgCO2e)" 的索引位置
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
                        text_value = carbon_emissions_element.text.strip()
                        # carbon_emissions = round(float(carbon_emissions_element.text),4)
                        carbon_emissions = Decimal(text_value).quantize(Decimal('0.0001'),rounding=ROUND_HALF_UP)
                        total_carbon_emissions += carbon_emissions
            # 判斷是否有下一頁,有則點擊,無則停止
            try:
                next_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                    (By.XPATH, "//li[@title='下一頁']//button[@class='ant-pagination-item-link']")))
                ActionChains(driver).move_to_element(next_button).click().perform()
                time.sleep(2)

            except Exception as e:
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

    # 檢查是否找到 "總計 kgCO2e" 的索引位置
    for i, cell in enumerate(header_cells):

        if cell.text == '總計 kgCO2e':
            emissions_index = i
    if emissions_index != -1:
        # 定位 <tbody> 元素
        tbody = table.find_element(By.TAG_NAME, 'tbody')

        # 定位所有的行元素
        rows = tbody.find_elements(By.TAG_NAME, 'tr')
        # 存放碳排放量加總的值
        total_emissions_elements = 0

        # 遍歷每個行元素
        row_end = rows[-1]
        emissions_elements = row_end.find_element(By.XPATH, f'.//td[{emissions_index + 1}]')
        emissions_element = round(float(emissions_elements.text),4)
        total_emissions_elements += emissions_element
        print(f'總計 kgCO2e : {total_emissions_elements}')
    else:
        print('未找到標題為 "總計 kgCO2e" 的列')

    return total_emissions_elements



# 點擊計算過程按鈕(冷媒)
def calculation_process_button_rfg(driver, i):
    carbon_emissions_dict = {}
    has_next_page = True
    while has_next_page:
        # 定位要滾動的table
        table_scroll_container = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//div[contains(@class, "ant-table-content")]'))
        )
        # 向右滾動到最底
        driver.execute_script("arguments[0].scrollLeft = arguments[0].scrollWidth;", table_scroll_container)
        time.sleep(2)
        # 定位table表單
        table = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
        thead_row = table.find_element(By.XPATH, './/thead/tr')

        # 找到 "碳排放量(kgCO2e)" 在表頭中的索引位置
        carbon_emissions_index = -1
        header_cells = thead_row.find_elements(By.TAG_NAME, 'th')
        for idx, cell in enumerate(header_cells):
            if cell.text == '碳排放量(kgCO2e)':
                carbon_emissions_index = idx
                break

        # 找到所有包含"計算過程"的按鈕
        buttons = table.find_elements(By.XPATH, './/td[@class="ant-table-cell calculate_table"]//span[text()="計算過程"]')
        if not buttons:
            print("找不到'計算過程'按鈕")
            return


        print("人為逸散 - 冷媒 計算過程 :")
        for index_2, button in enumerate(buttons):
            print(f"{i} - 第 {index_2 + 1} 筆 驗證:")
            # 找到與當前按鈕對應的"碳排放量(kgCO2e)"
            row = button.find_element(By.XPATH, "./ancestor::tr")  # 找到該按鈕所在的 `tr`
            carbon_emissions_element = row.find_element(By.XPATH, f'.//td[{carbon_emissions_index + 1}]')
            try:
                carbon_emissions = float(carbon_emissions_element.text)
                carbon_emissions_dict[index_2 + 1] = carbon_emissions
            except :
                print("無資料")
                continue
            # 點擊按鈕
            driver.execute_script("arguments[0].scrollIntoView();", button)  # 確保按鈕可見
            driver.execute_script("arguments[0].click();", button)  # 以 JavaScript 點擊，避免元素被遮擋
            time.sleep(2)

            # 執行計算結果驗證
            process_table_data_and_verify_rfg(driver,i,index_2,carbon_emissions_dict)
        try:
            next_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, "//li[@title='下一頁']//button[@class='ant-pagination-item-link']")))
            ActionChains(driver).move_to_element(next_button).click().perform()
            time.sleep(2)
        except Exception as e:
            has_next_page = False

    return carbon_emissions_dict

# 抓取排放係數、使用量、使用量*排放係數,並計算結果與驗證(移動,固定,製程)
def process_table_data_and_verify_rfg(driver,i,index_2,carbon_emissions_dict):
    table = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f'(//div[contains(@class, "ant-table-content")])[{index_2 + 2}]')))
    usage_index= get_column_index(table, "使用量")  # 找出"使用量"在table中的索引位置
    coefficient_index = get_column_index(table, "係數(設備逸散率)")  # 找出"係數(設備逸散率)"在table中的索引位置
    month_index = get_column_index(table, "使用月數")  # 找出"使用月數"在table中的索引位置
    verify_index = get_column_index(table,"使用量 X係數(設備逸散率) X使用月數 / 盤查月數")   # 找出"使用量 X係數(設備逸散率) X使用月數 / 盤查月數"在table中的索引位置
    gwp_index = get_column_index(table,"GWP")  # 找出"GWP"在table中的索引位置
    total_kgco2e_index = get_column_index(table,"排放量 kgCO2e")  # 找出"排放量 kgco2e"在table中的索引位置

    if -1 in [usage_index, coefficient_index, month_index, verify_index, gwp_index, total_kgco2e_index]:
        print("找不到元素")
        return
    # 定位table中的tbody/tr元素
    rows = table.find_elements(By.XPATH, './/tbody/tr[@class="ant-table-row ant-table-row-level-0"]')
    for row in rows:
        usage_elements = row.find_elements(By.XPATH, f'.//td[{usage_index + 1}]')
        coefficient_elements = row.find_elements(By.XPATH, f'.//td[{coefficient_index + 1}]')
        month_elements = row.find_elements(By.XPATH, f'.//td[{month_index + 1}]')
        verify_elements = row.find_elements(By.XPATH, f'.//td[{verify_index + 1}]')
        gwp_elements = row.find_elements(By.XPATH, f'.//td[{gwp_index + 1}]')
        total_kgco2e_elements = row.find_elements(By.XPATH, f'.//td[{total_kgco2e_index + 1}]')
        # 定位各欄位的值
        usage = float(usage_elements[0].text) if usage_elements else None
        coefficient = float(coefficient_elements[0].text) if coefficient_elements else None
        month = float(month_elements[0].text) if month_elements else None
        verify = float(verify_elements[0].text) if verify_elements else None
        gwp = float(gwp_elements[0].text) if gwp_elements else None
        total_kgco2e = float(total_kgco2e_elements[0].text) if total_kgco2e_elements else None
        calculated_results_1 = round(usage * month / 13, 4)
        calculated_results_2 = round(calculated_results_1 * coefficient, 4)
        calculated_results_gwp = round(calculated_results_2 * gwp, 4)
        total_carbon_emission = carbon_emissions_dict.get(index_2 + 1)
        if abs(calculated_results_2 - verify) < 1e-6:  # 容許極小誤差
            print(f"使用量 X係數(設備逸散率) X使用月數 / 盤查月數 : 程式計算結果: {calculated_results_2}, 頁面(計算過程)呈現: {verify} ✅ 一致")
        else:
            print(f"使用量 X係數(設備逸散率) X使用月數 / 盤查月數 : 程式計算結果: {calculated_results_2}, 頁面(計算過程)呈現: {verify} ❌ 不一致!!!")

        if abs(calculated_results_gwp - total_kgco2e) < 1e-6:
            print(f"使用量 X係數(設備逸散率) X使用月數 / 盤查月數 X gwp - 程式計算結果: {calculated_results_gwp}, 頁面(計算過程)呈現: {total_kgco2e} ✅ 一致")
        else:
            print(f"使用量 X係數(設備逸散率) X使用月數 / 盤查月數 X gwp - 程式計算結果: {calculated_results_gwp}, 頁面(計算過程)呈現: {total_kgco2e}  ❌ 不一致!!!")
        if abs(total_kgco2e - total_carbon_emission) < 1e-6:
            print(f"計算過程內總計 : {total_kgco2e}, 單筆碳排放量 :{total_carbon_emission} 一致 ✅")
        else :
            print(f"計算過程內總計 : {total_kgco2e}, 單筆碳排放量 :{total_carbon_emission} 不一致 ❌")

    # 當驗證完成後,點擊取消,關閉視窗
    try:
        modals = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, '//div[@class="ant-modal-content"]'))
        )
        latest_modal = modals[-1]  #取最後一個出現的 modal

        # 在這個 modal 內找到取消按鈕
        cancel_button = WebDriverWait(latest_modal, 10).until(
            EC.element_to_be_clickable((By.XPATH, './/button[@class="ant-btn ant-btn-default"]//span[text()="取 消"]'))
        )
        cancel_button.click()
        time.sleep(2)
    except Exception as e:
        print(f"無點擊到取消")

def calculate_co2_visitor(driver):
    total_visitor_emissions = Decimal('0.0000')
    has_next_page = True
    while has_next_page:
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
                        text_value = carbon_emissions_element.text.strip()
                        carbon_emissions = Decimal(text_value).quantize(Decimal('0.0001'),rounding=ROUND_HALF_UP)
                        total_visitor_emissions += carbon_emissions
            try:
                next_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable(
                    (By.XPATH, "//li[@title='下一頁']//button[@class='ant-pagination-item-link']")))
                ActionChains(driver).move_to_element(next_button).click().perform()
                time.sleep(2)

            except Exception:
                print(f'碳排放量總和：{total_visitor_emissions}')
                has_next_page = False
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

        if cell.text == '總計 kgCO2e':
            emissions_index = i

    # 檢查是否找到 "總計 kgCO2e" 的索引位置
    if emissions_index != -1:
        tbody = table.find_element(By.TAG_NAME, 'tbody')  #定位 <tbody> 元素
        rows = tbody.find_elements(By.TAG_NAME, 'tr')    #定位所有的行元素
        total_emissions_elements = 0

        row_end = rows[-1]  #遍歷每個行元素
        emissions_elements = row_end.find_element(By.XPATH, f'.//td[{emissions_index + 1}]')
        emissions_element = float(emissions_elements.text)
        total_emissions_elements += emissions_element
        print(f'排放量 KgCO2e : {total_emissions_elements}')
    else:
        print('未找到標題為 "排放量 KgCO2e" 的列')

    return total_emissions_elements

# 點擊計算過程按鈕(消防)
def calculation_process_button_fir(driver, i):
    carbon_emissions_dict = {}
    has_next_page = True
    while has_next_page:
        # 定位要滾動的table
        table_scroll_container = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//div[contains(@class, "ant-table-content")]'))
        )
        # 向右滾動到最底
        driver.execute_script("arguments[0].scrollLeft = arguments[0].scrollWidth;", table_scroll_container)
        time.sleep(2)
        # 定位table表單
        table = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
        thead_row = table.find_element(By.XPATH, './/thead/tr')

        # 找到 "碳排放量(kgCO2e)" 在表頭中的索引位置
        carbon_emissions_index = -1
        header_cells = thead_row.find_elements(By.TAG_NAME, 'th')
        for idx, cell in enumerate(header_cells):
            if cell.text == '碳排放量(kgCO2e)':
                carbon_emissions_index = idx
                break

        # 找到所有包含"計算過程"的按鈕
        buttons = table.find_elements(By.XPATH, './/td[@class="ant-table-cell calculate_table"]//span[text()="計算過程"]')
        if not buttons:
            print("找不到'計算過程'按鈕")
            return

        print("人為逸散 - 消防 計算結果 :")
        for index_2, button in enumerate(buttons):
            print(f"{i} - 第 {index_2 + 1} 筆 驗證:")

            # 找到與當前按鈕對應的"碳排放量(kgCO2e)"
            row = button.find_element(By.XPATH, "./ancestor::tr")  # 找到該按鈕所在的 `tr`
            carbon_emissions_element = row.find_element(By.XPATH, f'.//td[{carbon_emissions_index + 1}]')

            try:
                carbon_emissions = float(carbon_emissions_element.text)
                carbon_emissions_dict[index_2 + 1] = carbon_emissions
            except :
                print("無資料")
                continue
            # 點擊按鈕
            driver.execute_script("arguments[0].scrollIntoView();", button)  # 確保按鈕可見
            driver.execute_script("arguments[0].click();", button)  # 以 JavaScript 點擊，避免元素被遮擋
            time.sleep(2)

            # 執行計算結果驗證
            process_table_data_and_verify_fir(driver,i,index_2,carbon_emissions_dict)
        try:
            next_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, "//li[@title='下一頁']//button[@class='ant-pagination-item-link']")))
            ActionChains(driver).move_to_element(next_button).click().perform()
            time.sleep(2)
        except Exception as e:
            has_next_page = False

    return carbon_emissions_dict

# 抓取排放係數、使用量、使用量*排放係數,並計算結果與驗證(移動,固定,製程)
def process_table_data_and_verify_fir(driver,i,index_2,carbon_emissions_dict):
    table = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f'(//div[contains(@class, "ant-table-content")])[{index_2 + 2}]')))
    usage_index= get_column_index(table, "使用量")
    gwp_index = get_column_index(table,"GWP")
    total_kgco2e_index = get_column_index(table,"排放量 kgCO2e")
    if -1 in [usage_index, gwp_index, total_kgco2e_index]:
        print("找不到元素")
        return
    # 定位table中的tbody/tr元素
    rows = table.find_elements(By.XPATH, './/tbody/tr[@class="ant-table-row ant-table-row-level-0"]')
    for row in rows:
        usage_elements = row.find_elements(By.XPATH, f'.//td[{usage_index + 1}]')
        gwp_elements = row.find_elements(By.XPATH, f'.//td[{gwp_index + 1}]')
        total_kgco2e_elements = row.find_elements(By.XPATH, f'.//td[{total_kgco2e_index + 1}]')
        # 定位各欄位的值
        usage = float(usage_elements[0].text) if usage_elements else None
        gwp = float(gwp_elements[0].text) if gwp_elements else None
        total_kgco2e = float(total_kgco2e_elements[0].text) if total_kgco2e_elements else None
        calculated_results = round(usage * gwp, 4)  # 四捨五入(使用量 X GWP ,4)
        total_carbon_emission = carbon_emissions_dict.get(index_2 + 1)
        if abs(calculated_results - total_kgco2e) < 1e-6:  # 容許極小誤差
            print(f"使用量 X GWP : 程式計算結果: {calculated_results}, 頁面(計算過程)呈現: {total_kgco2e} 一致 ✅")
        else:
            print(f"使用量 X GWP : 程式計算結果: {calculated_results}, 頁面(計算過程)呈現: {total_kgco2e} 不一致!!! ❌")

        if abs(total_kgco2e - total_carbon_emission) < 1e-6:
            print(f"計算過程內總計 : {total_kgco2e}, 單筆碳排放量 :{total_carbon_emission} 一致 ✅")
        else :
            print(f"計算過程內總計 : {total_kgco2e}, 單筆碳排放量 :{total_carbon_emission} 不一致 ❌")

    # 當驗證完成後,點擊取消,關閉視窗
    try:
        modals = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, '//div[@class="ant-modal-content"]'))
        )
        latest_modal = modals[-1]  #取最後一個出現的 modal

        # 在這個 modal 內找到取消按鈕
        cancel_button = WebDriverWait(latest_modal, 10).until(
            EC.element_to_be_clickable((By.XPATH, './/button[@class="ant-btn ant-btn-default"]//span[text()="取 消"]'))
        )
        cancel_button.click()
        time.sleep(2)
    except Exception as e:
        print(f"無點擊到取消")

# 點擊計算過程按鈕(其他關注類物質 C1)
def calculation_process_button_other(driver, i):
    carbon_emissions_dict = {}
    has_next_page = True
    while has_next_page:
        # 定位要滾動的table
        table_scroll_container = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//div[contains(@class, "ant-table-content")]'))
        )
        # 向右滾動到最底
        driver.execute_script("arguments[0].scrollLeft = arguments[0].scrollWidth;", table_scroll_container)
        time.sleep(2)
        # 定位table表單
        table = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
        thead_row = table.find_element(By.XPATH, './/thead/tr')

        # 找到 "碳排放量(kgCO2e)" 在表頭中的索引位置
        carbon_emissions_index = -1
        header_cells = thead_row.find_elements(By.TAG_NAME, 'th')
        for idx, cell in enumerate(header_cells):
            if cell.text == '碳排放量(kgCO2e)':
                carbon_emissions_index = idx
                break

        # 找到所有包含"計算過程"的按鈕
        buttons = table.find_elements(By.XPATH, './/td[@class="ant-table-cell calculate_table"]//span[text()="計算過程"]')
        if not buttons:
            print("找不到'計算過程'按鈕")
            return
        print(f"{i} 計算結果 :")
        for index_2, button in enumerate(buttons):
            print(f"{i} - 第 {index_2 + 1} 筆 驗證:")

            # 找到與當前按鈕對應的"碳排放量(kgCO2e)"
            row = button.find_element(By.XPATH, "./ancestor::tr")  # 找到該按鈕所在的 `tr`
            carbon_emissions_element = row.find_element(By.XPATH, f'.//td[{carbon_emissions_index + 1}]')
            try:
                carbon_emissions = float(carbon_emissions_element.text)
                carbon_emissions_dict[index_2 + 1] = carbon_emissions
            except :
                print("無資料")
                continue
            # 點擊按鈕
            driver.execute_script("arguments[0].scrollIntoView();", button)  # 確保按鈕可見
            driver.execute_script("arguments[0].click();", button)  # 以 JavaScript 點擊，避免元素被遮擋
            time.sleep(2)

            # 執行計算結果驗證
            process_table_data_and_verify_other(driver,i,index_2,carbon_emissions_dict)
        try:
            next_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, "//li[@title='下一頁']//button[@class='ant-pagination-item-link']")))
            ActionChains(driver).move_to_element(next_button).click().perform()
            time.sleep(2)
        except Exception as e:
            has_next_page = False

    return carbon_emissions_dict

# 抓取排放係數、使用量、使用量*排放係數,並計算結果與驗證(移動,固定,製程)
def process_table_data_and_verify_other(driver,i,index_2,carbon_emissions_dict):
    table = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f'(//div[contains(@class, "ant-table-content")])[{index_2 + 2}]')))
    usage_index= get_column_index(table, "使用量")
    coefficient_index = get_column_index(table, "係數值")
    verify_index = get_column_index(table,"使用量 X係數值")
    gwp_index = get_column_index(table,"GWP")
    total_kgco2e_index = get_column_index(table,"排放量 kgCO2e")

    if -1 in [usage_index, coefficient_index, verify_index, gwp_index, total_kgco2e_index]:
        print("找不到元素")
        return
    # 定位table中的tbody/tr元素
    rows = table.find_elements(By.XPATH, './/tbody/tr[@class="ant-table-row ant-table-row-level-0"]')
    for row in rows:
        usage_elements = row.find_elements(By.XPATH, f'.//td[{usage_index + 1}]')
        coefficient_elements = row.find_elements(By.XPATH, f'.//td[{coefficient_index + 1}]')
        verify_elements = row.find_elements(By.XPATH, f'.//td[{verify_index + 1}]')
        gwp_elements = row.find_elements(By.XPATH, f'.//td[{gwp_index + 1}]')
        total_kgco2e_elements = row.find_elements(By.XPATH, f'.//td[{total_kgco2e_index + 1}]')
        # 定位各欄位的值
        usage = float(usage_elements[0].text) if usage_elements else None
        coefficient = float(coefficient_elements[0].text) if coefficient_elements else None
        verify = float(verify_elements[0].text) if verify_elements else None
        gwp = float(gwp_elements[0].text) if gwp_elements else None
        total_kgco2e = float(total_kgco2e_elements[0].text) if total_kgco2e_elements else None
        calculated_results = round(usage * coefficient, 4)
        calculated_results_gwp = round(calculated_results * gwp, 4)
        total_carbon_emission = carbon_emissions_dict.get(index_2 + 1)
        if abs(calculated_results - verify) < 1e-6:  # 容許極小誤差
            print(f"使用量 X係數值 : 程式計算結果: {calculated_results}, 頁面(計算過程)呈現: {verify} 一致 ✅")
        else:
            print(f"使用量 X係數值 : 程式計算結果: {calculated_results}, 頁面(計算過程)呈現: {verify} 不一致!!! ❌")

        if abs(calculated_results_gwp - total_kgco2e) < 1e-6:
            print(f"使用量 X係數(設備逸散率) X使用月數 / 盤查月數 X gwp - 程式計算結果: {calculated_results_gwp}, 頁面(計算過程)呈現: {total_kgco2e} ✅ 一致")
        else:
            print(f"使用量 X係數(設備逸散率) X使用月數 / 盤查月數 X gwp - 程式計算結果: {calculated_results_gwp}, 頁面(計算過程)呈現: {total_kgco2e}  ❌ 不一致!!!")
        if abs(total_kgco2e - total_carbon_emission) < 1e-6:
            print(f"計算過程內總計 : {total_kgco2e}, 單筆碳排放量 :{total_carbon_emission} 一致 ✅")
        else :
            print(f"計算過程內總計 : {total_kgco2e}, 單筆碳排放量 :{total_carbon_emission} 不一致 ❌")

    # 當驗證完成後,點擊取消,關閉視窗
    try:
        modals = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, '//div[@class="ant-modal-content"]'))
        )
        latest_modal = modals[-1]  #取最後一個出現的 modal

        # 在這個 modal 內找到取消按鈕
        cancel_button = WebDriverWait(latest_modal, 10).until(
            EC.element_to_be_clickable((By.XPATH, './/button[@class="ant-btn ant-btn-default"]//span[text()="取 消"]'))
        )
        cancel_button.click()
        time.sleep(2)
    except Exception as e:
        print(f"無點擊到取消")

# 點擊計算過程按鈕(C3~C6)
def calculation_process_button_c3(driver, i,upelec_upother):
    carbon_emissions_dict = {}
    has_next_page = True
    while has_next_page:
        # 定位要滾動的table
        table_scroll_container = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//div[contains(@class, "ant-table-content")]'))
        )
        # 向右滾動到最底
        driver.execute_script("arguments[0].scrollLeft = arguments[0].scrollWidth;", table_scroll_container)
        time.sleep(2)
        # 定位table表單
        table = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'table')))
        thead_row = table.find_element(By.XPATH, './/thead/tr')

        # 找到 "碳排放量(kgCO2e)" 在表頭中的索引位置
        carbon_emissions_index = -1
        header_cells = thead_row.find_elements(By.TAG_NAME, 'th')
        for idx, cell in enumerate(header_cells):
            if cell.text == '碳排放量(kgCO2e)':
                carbon_emissions_index = idx
                break

        # 找到所有包含"計算過程"的按鈕
        buttons = table.find_elements(By.XPATH, './/td[@class="ant-table-cell calculate_table"]//span[text()="計算過程"]')
        if not buttons:
            print("找不到'計算過程'按鈕")
            return
        if upelec_upother == 1 :
            print("輸入電力上游 計算結果 :")
        elif upelec_upother == 2 :
            print("輸入其他能源上游 計算結果")
        else:
            print(f"{i} 計算結果 :")
        for index_2, button in enumerate(buttons):
            print(f"{i} - 第 {index_2 + 1} 筆 驗證:")

            # 找到與當前按鈕對應的"碳排放量(kgCO2e)"
            row = button.find_element(By.XPATH, "./ancestor::tr")  # 找到該按鈕所在的 `tr`
            carbon_emissions_element = row.find_element(By.XPATH, f'.//td[{carbon_emissions_index + 1}]')

            try:
                carbon_emissions = float(carbon_emissions_element.text)
                carbon_emissions_dict[index_2 + 1] = carbon_emissions
            except :
                print("無資料")
                continue
            # 點擊按鈕
            driver.execute_script("arguments[0].scrollIntoView();", button)  # 確保按鈕可見
            driver.execute_script("arguments[0].click();", button)  # 以 JavaScript 點擊，避免元素被遮擋
            time.sleep(2)

            # 執行計算結果驗證
            process_table_data_and_verify_c3(driver,i,index_2,carbon_emissions_dict)
        try:
            next_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                (By.XPATH, "//li[@title='下一頁']//button[@class='ant-pagination-item-link']")))
            ActionChains(driver).move_to_element(next_button).click().perform()
            time.sleep(2)
        except Exception as e:
            has_next_page = False

    return carbon_emissions_dict

# 抓取排放係數、使用量、使用量*排放係數,並計算結果與驗證(移動,固定,製程)
def process_table_data_and_verify_c3(driver,i,index_2,carbon_emissions_dict):
    table = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f'(//div[contains(@class, "ant-table-content")])[{index_2 + 2}]')))
    usage_index= get_column_index(table, "活動強度(使用量)")
    coefficient_index = get_column_index(table, "係數值")
    total_kgco2e_index = get_column_index(table,"排放量 kgCO2e")

    if -1 in [usage_index, coefficient_index, total_kgco2e_index]:
        print("找不到元素")
        return
    # 定位table中的tbody/tr元素
    rows = table.find_elements(By.XPATH, './/tbody/tr[@class="ant-table-row ant-table-row-level-0"]')
    for row in rows:
        usage_elements = row.find_elements(By.XPATH, f'.//td[{usage_index + 1}]')
        coefficient_elements = row.find_elements(By.XPATH, f'.//td[{coefficient_index + 1}]')
        total_kgco2e_elements = row.find_elements(By.XPATH, f'.//td[{total_kgco2e_index + 1}]')
        # 定位各欄位的值
        usage = float(usage_elements[0].text) if usage_elements else None
        coefficient = float(coefficient_elements[0].text) if coefficient_elements else None
        total_kgco2e = float(total_kgco2e_elements[0].text) if total_kgco2e_elements else None
        calculated_results = round(usage * coefficient, 4)
        total_carbon_emission = carbon_emissions_dict.get(index_2 + 1)
        if abs(calculated_results - total_kgco2e) < 1e-6:  # 容許極小誤差
            print(f"使用量 X係數值 : 程式計算結果: {calculated_results}, 頁面(計算過程)呈現: {total_kgco2e} 一致✅")
        else:
            print(f"使用量 X係數值 : 程式計算結果: {calculated_results}, 頁面(計算過程)呈現: {total_kgco2e} 不一致!!! ❌")
        if abs(total_kgco2e - total_carbon_emission) < 1e-6:
            print(f"計算過程內總計 : {total_kgco2e}, 單筆碳排放量 :{total_carbon_emission} 一致 ✅")
        else :
            print(f"計算過程內總計 : {total_kgco2e}, 單筆碳排放量 :{total_carbon_emission} 不一致 ❌")

    # 當驗證完成後,點擊取消,關閉視窗
    try:
        modals = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, '//div[@class="ant-modal-content"]'))
        )
        latest_modal = modals[-1]  #取最後一個出現的 modal

        # 在這個 modal 內找到取消按鈕
        cancel_button = WebDriverWait(latest_modal, 10).until(
            EC.element_to_be_clickable((By.XPATH, './/button[@class="ant-btn ant-btn-default"]//span[text()="取 消"]'))
        )
        cancel_button.click()
        time.sleep(2)
    except Exception as e:
        print(f"無點擊到取消")


# 總值驗證結果(C3~C6)
def validate_co2_visitor_c3(driver):
    table = driver.find_element(By.XPATH, '//table[contains(@class, "table table-sm table-bordered text-end")]')
    thead_row = table.find_element(By.XPATH, './/thead/tr')

    # 找到 "排放源" 和 "排放量 kgCO2e" 在表頭中的索引位置
    emissions_index = -1
    header_cells = thead_row.find_elements(By.TAG_NAME, 'th')
    for i, cell in enumerate(header_cells):
        if cell.text == '總計 kgCO2e':
            emissions_index = i

    # 檢查是否找到 "總計 kgCO2e" 的索引位置
    if emissions_index != -1:
        tbody = table.find_element(By.TAG_NAME, 'tbody')  #定位 <tbody> 元素
        rows = tbody.find_elements(By.TAG_NAME, 'tr')    #定位所有的行元素
        total_emissions_elements = 0

        row_end = rows[-1]  #遍歷每個行元素
        emissions_elements = row_end.find_element(By.XPATH, f'.//td[{emissions_index}]')
        emissions_element = round(float(emissions_elements.text),4)
        total_emissions_elements += emissions_element
        print(f'排放量 KgCO2e : {total_emissions_elements}')
    else:
        print('未找到標題為 "排放量 KgCO2e" 的列')

    return total_emissions_elements

def calculation_process_button_workhour(driver, i):
    carbon_emissions_dict = {}
    # 定位table表單
    table = driver.find_element(By.TAG_NAME, 'table')
    # 定位表頭行元素
    tbody_row = table.find_element(By.TAG_NAME, 'tbody')
    rows = tbody_row.find_elements(By.TAG_NAME, 'tr')
    carbon_emissions_index = None
    for index, row in enumerate(rows):
        first_cell = row.find_element(By.XPATH, './td[1]')
        if first_cell.text.strip() == '碳排放量(kgCO2e)':
            carbon_emissions_index = index  # 存 index，而不是 WebElement
            break
    # 找到所有包含"計算過程"的按鈕
    buttons = table.find_elements(By.XPATH, './/td[@class="ant-table-cell calculate_table"]//span[text()="計算過程"]')
    if not buttons:
        print("找不到'計算過程'按鈕")
        return
    # 存取每筆碳排放量
    print(f"{i} 計算結果 :")
    for index_2, button in enumerate(buttons):
        print(f"{i} - 第 {index_2 + 1} 筆 驗證:")
        # 找到與當前按鈕對應的"碳排放量(kgCO2e)"
        row = button.find_element(By.XPATH, "./ancestor::tr")  # 找到該按鈕所在的 `tr`
        next_row = row.find_element(By.XPATH, "./following-sibling::tr[1]")
        carbon_emissions_element = next_row.find_elements(By.TAG_NAME, "td")[index_2 + 1]  # 定位該筆的碳排放量
        # carbon_emissions = float(carbon_emissions_element.text)
        try:
            carbon_emissions = round(Decimal(carbon_emissions_element.text),4)
            carbon_emissions_dict[index_2 + 1] = carbon_emissions
        except :
            print("無資料")
            continue
        # 點擊按鈕
        driver.execute_script("arguments[0].scrollIntoView();", button)  # 確保按鈕可見
        driver.execute_script("arguments[0].click();", button)  # 以 JavaScript 點擊，避免元素被遮擋
        time.sleep(2)

        # 執行計算結果驗證
        process_table_data_and_verify_workhour(driver, i,index_2,carbon_emissions_dict)

    return carbon_emissions_dict

# 抓取排放係數、使用量、使用量*排放係數,並計算結果與驗證(移動,固定,製程)
def process_table_data_and_verify_workhour(driver,i,index_2,carbon_emissions_dict):
    table = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, f'(//div[contains(@class, "ant-table-content")])[{index_2 + 2}]')))
    usage_index = get_column_index(table, "總工時")  # 找出"使用量"在table中的索引位置
    coefficient_index = get_column_index(table, "CH4排放係數(kg/人時)")  # 找出"係數(設備逸散率)"在table中的索引位置
    verify_index = get_column_index(table, "總工時 X 係數")  # 找出"使用量 X係數(設備逸散率) X使用月數 / 盤查月數"在table中的索引位置
    gwp_index = get_column_index(table, "GWP")  # 找出"GWP"在table中的索引位置
    total_kgco2e_index = get_column_index(table, "排放量 kgCO2e")  # 找出"排放量 kgco2e"在table中的索引位置

    if -1 in [usage_index, coefficient_index, verify_index, gwp_index, total_kgco2e_index]:
        print("找不到元素")
        return
    # 定位table中的tbody/tr元素
    row = table.find_element(By.XPATH, './/tbody/tr[@class="ant-table-row ant-table-row-level-0"]')

    usage_elements = row.find_elements(By.XPATH, f'.//td[{usage_index + 1}]')
    coefficient_elements = row.find_elements(By.XPATH, f'.//td[{coefficient_index + 1}]')
    verify_elements = row.find_elements(By.XPATH, f'.//td[{verify_index + 1}]')
    gwp_elements = row.find_elements(By.XPATH, f'.//td[{gwp_index + 1}]')
    total_kgco2e_elements = row.find_elements(By.XPATH, f'.//td[{total_kgco2e_index + 1}]')
    # 定位各欄位的值
    usage = Decimal(usage_elements[0].text) if usage_elements else None
    coefficient = Decimal(coefficient_elements[0].text) if coefficient_elements else None
    verify = Decimal(verify_elements[0].text) if verify_elements else None
    gwp = Decimal(gwp_elements[0].text) if gwp_elements else None
    total_kgco2e = Decimal(total_kgco2e_elements[0].text) if total_kgco2e_elements else None
    calculated_results = Decimal(usage * coefficient).quantize(Decimal("0.0001"), rounding = ROUND_HALF_UP)
    calculated_results_gwp = Decimal(calculated_results * gwp).quantize(Decimal("0.0001"),rounding = ROUND_HALF_UP)
    total_carbon_emission = carbon_emissions_dict.get(index_2 + 1)
    if abs(calculated_results - verify) < 1e-6:  # 容許極小誤差
        print(
            f"總工時 X 係數 : 程式計算結果: {calculated_results}, 頁面(計算過程)呈現: {verify} ✅ 一致")
    else:
        print(
            f"總工時 X 係數 : 程式計算結果: {calculated_results}, 頁面(計算過程)呈現: {verify} ❌ 不一致!!!")

    if abs(calculated_results_gwp - total_kgco2e) < 1e-6:
        print(
            f"總工時 X 係數 X gwp - 程式計算結果: {calculated_results_gwp}, 頁面(計算過程)呈現: {total_kgco2e} ✅ 一致")
    else:
        print(
            f"總工時 X 係數 X gwp - 程式計算結果: {calculated_results_gwp}, 頁面(計算過程)呈現: {total_kgco2e}  ❌ 不一致!!!")
    if abs(total_kgco2e - total_carbon_emission) < 1e-6:
        print(f"計算過程內總計 : {total_kgco2e}, 單筆碳排放量 :{total_carbon_emission} 一致 ✅")
    else:
        print(f"計算過程內總計 : {total_kgco2e}, 單筆碳排放量 :{total_carbon_emission} 不一致 ❌")

    # 當驗證完成後,點擊取消,關閉視窗
    try:
        modals = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, '//div[@class="ant-modal-content"]'))
        )
        latest_modal = modals[-1]  # 取最後一個出現的 modal

        # 在這個 modal 內找到取消按鈕
        cancel_button = WebDriverWait(latest_modal, 10).until(
            EC.element_to_be_clickable((By.XPATH, './/button[@class="ant-btn ant-btn-default"]//span[text()="取 消"]'))
        )
        cancel_button.click()
        time.sleep(2)
    except Exception as e:
        print(f"無點擊到取消")

def calculate_co2_workhour(driver):
    total_visitor_emissions = Decimal('0.0000')
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
            text_value = carbon_emissions_element.text.strip()
            carbon_emissions = Decimal(text_value).quantize(Decimal('0.0001'),rounding=ROUND_HALF_UP)

            total_visitor_emissions += carbon_emissions

        # 更新燃料種類的碳排放量總和
        # print(f'碳排放量總和：{total_visitor_emissions}')
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
        if cell.text == '總計 kgCO2e':
            emissions_index = i

    # 檢查是否找到 "排放源" 和 "排放量 kgCO2e" 的索引位置
    if emissions_index != -1:
        # 定位 <tbody> 元素
        tbody = table.find_element(By.TAG_NAME, 'tbody')
        # 定位所有的行元素
        rows = tbody.find_elements(By.TAG_NAME, 'tr')
        total_emissions_elements = 0

        # 遍歷每個行元素
        row_end = rows[-1]
        emissions_elements = row_end.find_element(By.XPATH, f'.//td[{emissions_index} ]')
        emissions_element = float(emissions_elements.text)

        total_emissions_elements += emissions_element
        # 更新排放源的排放量總和
        print(f'總計 kgCO2e : {total_emissions_elements}')
    else:
        print('未找到標題為 "總計 kgCO2e" 的列')

    return total_emissions_elements

def calculation_process_button_elec(driver, i):
    carbon_emissions_dict = {}
    # 定位table表單
    table = driver.find_element(By.TAG_NAME, 'table')
    # 定位表頭行元素
    tbody_row = table.find_element(By.TAG_NAME, 'tbody')
    rows = tbody_row.find_elements(By.TAG_NAME, 'tr')
    carbon_emissions_index = None
    for index, row in enumerate(rows):
        first_cell = row.find_element(By.XPATH, './td[1]')
        if first_cell.text.strip() == '碳排放量(kgCO2e)':
            carbon_emissions_index = index  # 存 index，而不是 WebElement
            break
    # 找到所有包含"計算過程"的按鈕
    buttons = table.find_elements(By.XPATH, './/td[@class="ant-table-cell calculate_table"]//span[text()="計算過程"]')
    if not buttons:
        print("找不到'計算過程'按鈕")
        return
    print(f"{i} 計算結果 :")
    for index_2, button in enumerate(buttons):
        print(f"{i} - 第 {index_2 + 1} 筆 驗證:")
        # 找到與當前按鈕對應的"碳排放量(kgCO2e)"
        row = button.find_element(By.XPATH, "./ancestor::tr")  # 找到該按鈕所在的 `tr`
        next_row = row.find_element(By.XPATH, "./following-sibling::tr[1]")
        carbon_emissions_element = next_row.find_elements(By.TAG_NAME, "td")[index_2 + 2]  # 定位該筆的碳排放量
        # carbon_emissions = float(carbon_emissions_element.text)
        try:
            carbon_emissions = float(carbon_emissions_element.text)
            carbon_emissions_dict[index_2 + 1] = carbon_emissions
        except :
            print("無資料")
            continue
        # 點擊按鈕
        driver.execute_script("arguments[0].scrollIntoView();", button)  # 確保按鈕可見
        driver.execute_script("arguments[0].click();", button)  # 以 JavaScript 點擊，避免元素被遮擋
        time.sleep(2)

        # 執行計算結果驗證
        process_table_data_and_verify_elec(driver, i,index_2,carbon_emissions_dict)

    return carbon_emissions_dict

# 抓取排放係數、使用量、使用量*排放係數,並計算結果與驗證(移動,固定,製程)
def process_table_data_and_verify_elec(driver,i,index_2,carbon_emissions_dict):
    table = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, f'(//div[contains(@class, "ant-table-content")])[{index_2 + 2}]')))
    usage_index = get_column_index(table, "總用電量")  # 找出"使用量"在table中的索引位置
    coefficient_index = get_column_index(table, "係數值")  # 找出"係數(設備逸散率)"在table中的索引位置
    total_kgco2e_index = get_column_index(table, "排放量 kgCO2e")  # 找出"排放量 kgco2e"在table中的索引位置

    if -1 in [usage_index, coefficient_index, total_kgco2e_index]:
        print("找不到元素")
        return
    # 定位table中的tbody/tr元素
    row = table.find_element(By.XPATH, './/tbody/tr[@class="ant-table-row ant-table-row-level-0"]')

    usage_elements = row.find_elements(By.XPATH, f'.//td[{usage_index + 1}]')
    coefficient_elements = row.find_elements(By.XPATH, f'.//td[{coefficient_index + 1}]')
    total_kgco2e_elements = row.find_elements(By.XPATH, f'.//td[{total_kgco2e_index + 1}]')
    # 定位各欄位的值
    usage = float(usage_elements[0].text) if usage_elements else None
    coefficient = float(coefficient_elements[0].text) if coefficient_elements else None
    total_kgco2e = float(total_kgco2e_elements[0].text) if total_kgco2e_elements else None
    calculated_results = round(usage * coefficient, 4)
    total_carbon_emission = carbon_emissions_dict.get(index_2 + 1)
    if abs(calculated_results - total_kgco2e) < 1e-6:
        print(
            f"總用電量 X 係數 - 程式計算結果: {calculated_results}, 頁面(計算過程)呈現: {total_kgco2e} ✅ 一致")
    else:
        print(
            f"總用電量 X 係數 - 程式計算結果: {calculated_results}, 頁面(計算過程)呈現: {total_kgco2e}  ❌ 不一致!!!")
    if abs(total_kgco2e - total_carbon_emission) < 1e-6:
        print(f"計算過程內總計 : {total_kgco2e}, 單筆碳排放量 :{total_carbon_emission} 一致 ✅")
    else:
        print(f"計算過程內總計 : {total_kgco2e}, 單筆碳排放量 :{total_carbon_emission} 不一致 ❌")

    # 當驗證完成後,點擊取消,關閉視窗
    try:
        modals = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, '//div[@class="ant-modal-content"]'))
        )
        latest_modal = modals[-1]  # 取最後一個出現的 modal

        # 在這個 modal 內找到取消按鈕
        cancel_button = WebDriverWait(latest_modal, 10).until(
            EC.element_to_be_clickable((By.XPATH, './/button[@class="ant-btn ant-btn-default"]//span[text()="取 消"]'))
        )
        cancel_button.click()
        time.sleep(2)
    except Exception as e:
        print(f"無點擊到取消")
