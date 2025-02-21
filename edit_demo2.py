import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from response import catch_response


def edit_demo2(driver, url):
    source_type = ['輸入能源上游排放 C4','移動源 C1','固定燃燒源 C1','工業製程 C1','人為逸散 C1','其他關注類物質 C1','輸入電力 C2','客戶和訪客運輸 C3','員工差旅 C3','員工通勤 C3','營運產生之廢棄物 C4','資本設備 C4','用水(水資源管理) B.5.2.a','購買商品 C4','上游租賃資產 C4','顧問諮詢、清潔、維護等 C4','上游運輸及配送 C3','廢棄物運輸 C3','下游運輸及配送 C3','下游租賃資產 C5','產品壽命終止階段 C5','產品使用階段 C5','投資 C5', '其它間接排放 C6']


    url = url+"/project/calculate/"

    for i in source_type:
        if i == '移動源 C1' or i == '固定燃燒源 C1' or i == '工業製程 C1' or i == '其他關注類物質 C1':
            link = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.LINK_TEXT, i)))
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))
            link.click()

            element = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, '//span[@aria-label="edit"]')))
            driver.execute_script("arguments[0].click();", element)
            input_data_C1_C2(driver, i, 0)

            time.sleep(2)

        elif i == '人為逸散 C1':
            link = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.LINK_TEXT, i)))
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))
            link.click()

            time.sleep(2)
            workhour_edit(driver, url)

            element = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//div[text()="冷媒設備 B.2.2.d"]')))
            element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//div[text()="冷媒設備 B.2.2.d"]')))
            driver.execute_script("arguments[0].click()", element)

            element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//span[@aria-label="edit"]')))
            driver.execute_script("arguments[0].click();", element)

            time.sleep(2)
            input_data_C1_C2(driver, i, 1)

            element = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//div[text()="消防設備 B.2.2.d"]')))
            element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//div[text()="消防設備 B.2.2.d"]')))
            driver.execute_script("arguments[0].click()", element)
            time.sleep(2)

            element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//span[@aria-label="edit"]')))
            driver.execute_script("arguments[0].click();", element)

            time.sleep(2)
            input_data_C1_C2(driver, i, 2)
        elif i == '輸入能源上游排放 C4':
            link = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.LINK_TEXT, i)))
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))
            link.click()

            element = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//div[text()="輸入電力上游 B.5.2.a"]')))
            element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//div[text()="輸入電力上游 B.5.2.a"]')))
            driver.execute_script("arguments[0].click()", element)

            element = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, '//span[@aria-label="edit"]')))
            driver.execute_script("arguments[0].click();", element)
            time.sleep(2)
            input_data_C4(driver, i, 1)

            element = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//div[text()="其他輸入能源上游 B.5.2.a"]')))
            element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//div[text()="其他輸入能源上游 B.5.2.a"]')))
            driver.execute_script("arguments[0].click()", element)
            time.sleep(2)
            element = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, '//span[@aria-label="edit"]')))
            driver.execute_script("arguments[0].click();", element)

            time.sleep(2)
            input_data_C4(driver, i, 2)
        elif i == '輸入電力 C2':
            link = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.LINK_TEXT, i)))
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))
            link.click()

            elec_edit(driver, url)
            time.sleep(2)

            # driver.get(url+'electricity')
            # element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//div[text()="綠電 B.3.2.a"]')))
            # element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[text()="綠電 B.3.2.a"]')))
            # driver.execute_script("arguments[0].click()", element)

            # time.sleep(3)
            # green_elec(driver,url)
        elif i == '輸入蒸汽 C2':
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))
            link.click()

            element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//div[text()="蒸氣加項 B.3.2.b"]')))
            driver.execute_script("arguments[0].click()", element)

            edit_steam_input(driver, url)
            time.sleep(3)
        else:
            link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, i)))
            link.click()

            element = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, '//span[@aria-label="edit"]')))
            driver.execute_script("arguments[0].click();", element)

            time.sleep(3)
            input_data_C3_C6(driver, i, 0)

        driver.get(url)


def input_data_C1_C2(driver, source_type, rfg_or_fire):
    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "validateOnly")))

    form_element = driver.find_element(By.ID, "validateOnly")  # 表單element

    input_elements = form_element.find_elements(By.TAG_NAME, "input")  # tag_name是input的

    CarPlateNo = ["ABC-789"]
    validateOnlyName1 = ["固定設備名稱測試編輯"]
    validateOnlyName2 = ["製程設備名稱測試編輯"]
    validateOnlyName3 = ["設備名稱測試編輯"]

    new_inputelements = []

    for i in input_elements:  # 把disable的先移除掉
        if i.get_attribute("disabled") != None and i.get_attribute(
                "id") != "validateOnly_referenceFile" and i.get_attribute("type") != "search":
            print("")
        else:
            new_inputelements.append(i)
    for input_element in new_inputelements:
        try:
            if input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute(
                    "id") == 'validateOnly_CarPlateNo':  # 車牌
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(CarPlateNo)
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute(
                    "id") == 'validateOnly_ProcessName':  # 製程別
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("製程別測試編輯")
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute(
                    "id") == 'validateOnly_ResponsibleUnit':  # 負責單位
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("負責單位測試編輯")
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute(
                    "id") == 'validateOnly_Name':  # 公務車型/設備名稱
                if source_type == "移動源 C1":
                    input_element.send_keys(Keys.CONTROL, 'a')
                    input_element.send_keys(Keys.BACKSPACE)
                    WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("公務車型測試編輯1")
                    time.sleep(0.5)
                elif source_type == "固定燃燒源 C1":
                    input_element.send_keys(Keys.CONTROL, 'a')
                    input_element.send_keys(Keys.BACKSPACE)
                    WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(validateOnlyName1)
                    time.sleep(0.5)
                elif source_type == "工業製程 C1":
                    input_element.send_keys(Keys.CONTROL, 'a')
                    input_element.send_keys(Keys.BACKSPACE)
                    WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(validateOnlyName2)
                    time.sleep(0.5)
                else:
                    input_element.send_keys(Keys.CONTROL, 'a')
                    input_element.send_keys(Keys.BACKSPACE)
                    WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(validateOnlyName3)
                    time.sleep(0.5)

            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute(
                    "id") == 'validateOnly_ActivityIntensityUnit':  # 使用量單位
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("KG")
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute(
                    "id") == 'validateOnly_ModelNumber':  # 使用量單位
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("冷媒型號測試編輯")
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute(
                    "id") == 'validateOnly_Evidence':  # 使用量佐證說明
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("使用量佐證說明測試編輯")
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute(
                    "id") == 'validateOnly_Description':  # 備註
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("備註測試編輯")
                time.sleep(1)

            elif input_element.get_attribute("id") == 'validateOnly_ingredientName':
                continue

            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute(
                    "id") == 'validateOnly_Scalar':  # 使用量/使用冷媒/製冷劑填充量
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(50)
                time.sleep(1)
            # elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_AnnualPurchaseAmount':
            #     WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys( ''.join(random.choices(string.digits, k=3)))
            #     time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute(
                    "id") == 'validateOnly_OriginalFill':  # 消防原始填充量
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(50)
                time.sleep(1)
            # elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute("id") == 'validateOnly_TransportDistance':
            #     WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys( ''.join(random.choices(string.digits, k=3)))
            #     time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute(
                    "id") == 'validateOnly_TotalNumber':
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(20)
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute(
                    "id") == 'validateOnly_UsedMonth':  # 使用月數
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(12)
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute(
                    "id") == 'validateOnly_Quantity':  # 消防數量/其他關注類數量
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(20)
                time.sleep(1)

            elif input_element.get_attribute("placeholder") == "請輸入活動強度單位":
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("KG")

            elif input_element.get_attribute("type") == "search":  # 選單類 先click 再選 再click

                if input_element.get_attribute("aria-activedescendant") == "validateOnly_Type_list_0":  # 車輛類別
                    element = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '公務車')]")))
                    element.click()

                elif input_element.get_attribute("id") == "validateOnly_Area":  # 地區
                    try:
                        element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                            (By.XPATH, "//span[contains(text(),'台灣 / 移動源 - 汽油KgCO2e/L / 2023')]")))  # 移動源係數
                        element.click()
                        element = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'台灣')]")))
                        element.click()
                        element = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'移動源 - 柴油')]")))  # 移動源係數
                        element.click()
                    except:
                        try:
                            element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                                (By.XPATH, "//span[contains(text(),'台灣 / 固定源 - 汽油 KgCO2e/L / 2023')]")))  # 移動源係數
                            element.click()
                            element = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'台灣')]")))
                            element.click()
                            element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                                (By.XPATH, "//div[contains(text(),'固定源 - 柴油 ')]")))  # 固定源係數
                            element.click()
                            time.sleep(0.5)
                        except:
                            element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                                (By.XPATH, "//span[contains(text(),'台灣 / 無煙煤 KgCO2e/Kg / 2023')]")))  # 移動源係數
                            element.click()
                            element = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'台灣')]")))
                            element.click()
                            element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                                (By.XPATH, "//div[contains(text(),'天然氣 KgCO2e/M3')]")))  # 製程係數
                            element.click()
                            time.sleep(0.5)
                    element = WebDriverWait(driver, 2).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'2023')]")))
                    element.click()
                    time.sleep(0.5)


                elif input_element.get_attribute(
                        "aria-activedescendant") == "validateOnly_activityDataType_list_0":  # 活動數據種類
                    element = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '自動連續量測')]")))
                    element.click()
                    element = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '間歇量測')]")))
                    element.click()
                    time.sleep(1)
                elif input_element.get_attribute(
                        "aria-activedescendant") == "validateOnly_emitParaType_list_0":  # 排放係數種類
                    element = WebDriverWait(driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '自我/量測/質能')]")))
                    element.click()
                    element = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '同製程/設備經驗')]")))
                    element.click()
                    time.sleep(1)
                elif input_element.get_attribute("id") == "validateOnly_ARnGWPid":
                    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                        (By.XPATH, "//span[contains(text(),'CO2(Carbon dioxide)')]")))  # 移動源係數
                    element.click()
                    element = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'CH4(Methane)')]")))
                    element.click()
                    time.sleep(1)
                elif input_element.get_attribute("id") == "validateOnly_Commuting":
                    input_element.click()
                    element1 = WebDriverWait(driver, 10).until(EC.visibility_of_element_located(
                        (By.XPATH, '//div[@class="ant-select-item-option-content"][2]')))
                    element1.click()
                    time.sleep(1)
                elif input_element.get_attribute("id") == "validateOnly_TransportTypeID":
                    input_element.click()
                    element1 = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH,
                                                                                                 '//div[@class="ant-select-item ant-select-item-option ant-pro-filed-search-select-option "][2]')))
                    element1.click()
                    time.sleep(1)
                elif input_element.get_attribute("id") == "validateOnly_warmGasType":
                    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                        (By.XPATH, "//span[contains(text(),'固定燃燒')]")))  # 移動源係數
                    element.click()
                    element = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '(逸散)其他逸散')]")))
                    element.click()
                    time.sleep(1)

                elif input_element.get_attribute("id") == "validateOnly_ParameterID":
                    try:
                        element = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(),'R-11')]")))  # 移動源係數
                        element.click()
                        element = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'R-114')]")))  # 使用冷媒/製冷劑種類
                        element.click()
                        time.sleep(1)
                    except:
                        element = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(),'R-123')]")))  # 移動源係數
                        element.click()
                        element = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Halon-1201')]")))
                        element.click()
                        time.sleep(1)
                elif input_element.get_attribute("id") == "validateOnly_ParameterID2":  # 設備類型(排放因子)
                    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                        (By.XPATH, "//span[contains(text(),'家用冰箱Household refrigerator')]")))  # 移動源係數
                    element.click()
                    element = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '冰水機Chiller')]")))
                    element.click()
                    time.sleep(1)
                else:
                    continue
                time.sleep(1)

        except:  # StaleElementReferenceException
            form_element = driver.find_element(By.ID, "validateOnly")  # 表單element
            input_elements = form_element.find_elements(By.TAG_NAME, "input")
            if source_type == "其他關注類物質 C1":
                WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.ID, "validateOnly_Unit"))).send_keys("t")
                time.sleep(1)
                WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.ID, "validateOnly_ActivityIntensityUnit"))).send_keys("Kg")
                time.sleep(1)

    button = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']")))  # 送出
    button.click()
    time.sleep(3)
    if rfg_or_fire == 1:
        catch_response(driver, '人為逸散 C1(冷媒設備)')
    elif rfg_or_fire == 2:
        catch_response(driver, '人為逸散 C1(消防設備)')
    elif rfg_or_fire == 3:
        catch_response(driver, '輸入能源上游排放 C4(輸入電力上游)')
    elif rfg_or_fire == 4:
        catch_response(driver, '輸入能源上游排放 C4(其他輸入能源上游)')
    elif rfg_or_fire == 0:
        catch_response(driver, source_type)
    time.sleep(2)


def wait_and_click(driver, by, value, wait_time=2):
    element = WebDriverWait(driver, wait_time).until(EC.element_to_be_clickable((by, value)))
    element.click()
    time.sleep(1)  # 若需要等待一段時間


def coefficient_C4_C6(driver):
    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "validateOnly")))
    time.sleep(2)
    wait_and_click(driver, By.XPATH,"//button[@class='ant-btn css-dts6b9 ant-btn-link d-flex align-items-center mt-2']")

    wait_and_click(driver, By.XPATH, "//div[@class='ant-collapse-header']//span[contains(text(), '環保署產品碳足跡資訊網20240624')]")

    wait_and_click(driver, By.XPATH, "//div[@class='ant-collapse-header']//span[contains(text(), '乳品及其加工品')]")

    wait_and_click(driver, By.XPATH, "//a[@class='d-block my-2 text-dark'][2]")

    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable(
        (By.XPATH, "//tr[@class='ant-table-row ant-table-row-level-0'][1]//span[@class='ant-checkbox-inner']")))
    actions = ActionChains(driver)
    actions.move_to_element(element).click().perform()
    time.sleep(1)

    wait_and_click(driver, By.XPATH,
                   "//button[@class='ant-btn css-dts6b9 ant-btn-primary']//span[contains(text(),'確 定')]")


def coefficient_C3(driver):
    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "validateOnly")))
    time.sleep(2)
    wait_and_click(driver, By.XPATH,
                   "//button[@class='ant-btn css-dts6b9 ant-btn-link d-flex align-items-center mt-2']")

    wait_and_click(driver, By.XPATH, "//div[@class='ant-collapse-header']//span[contains(text(), '環保署產品碳足跡資訊網20240624')]")

    wait_and_click(driver, By.XPATH, "//div[@class='ant-collapse-header']//span[contains(text(), '運輸服務')]")

    wait_and_click(driver, By.XPATH, "//a[@class='d-block my-2 text-dark'][2]")

    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable(
        (By.XPATH, "//tr[@class='ant-table-row ant-table-row-level-0'][1]//span[@class='ant-checkbox-inner']")))
    actions = ActionChains(driver)
    actions.move_to_element(element).click().perform()
    time.sleep(1)

    wait_and_click(driver, By.XPATH,
                   "//button[@class='ant-btn css-dts6b9 ant-btn-primary']//span[contains(text(),'確 定')]")


def coefficient_C4(driver):
    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "validateOnly")))
    time.sleep(2)
    wait_and_click(driver, By.XPATH,
                   "//button[@class='ant-btn css-dts6b9 ant-btn-link d-flex align-items-center mt-2']")

    wait_and_click(driver, By.XPATH, "//div[@class='ant-collapse-header']//span[contains(text(), '環保署產品碳足跡資訊網20240624')]")

    wait_and_click(driver, By.XPATH, "//div[@class='ant-collapse-header']//span[contains(text(), '廢棄物回收處理服務')]")

    wait_and_click(driver, By.XPATH, "//a[@class='d-block my-2 text-dark'][2]")

    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable(
        (By.XPATH, "//tr[@class='ant-table-row ant-table-row-level-0'][1]//span[@class='ant-checkbox-inner']")))
    actions = ActionChains(driver)
    actions.move_to_element(element).click().perform()
    time.sleep(1)

    wait_and_click(driver, By.XPATH,
                   "//button[@class='ant-btn css-dts6b9 ant-btn-primary']//span[contains(text(),'確 定')]")


def coefficient_C4_upstreamEmissions_elec(driver):
    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "validateOnly")))
    time.sleep(2)
    wait_and_click(driver, By.XPATH,
                   "//button[@class='ant-btn css-dts6b9 ant-btn-link d-flex align-items-center mt-2']")

    wait_and_click(driver, By.XPATH, "//div[@class='ant-collapse-item'][4]")

    wait_and_click(driver, By.XPATH, "//div[@class='ant-collapse-item']//span[contains(text(), '能資源')]")

    wait_and_click(driver, By.XPATH, "//a[@class='d-block my-2 text-dark'][6]")

    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable(
        (By.XPATH, "//tr[@class='ant-table-row ant-table-row-level-0'][2]//span[@class='ant-checkbox-inner']")))
    actions = ActionChains(driver)
    actions.move_to_element(element).click().perform()
    time.sleep(1)

    wait_and_click(driver, By.XPATH,
                   "//button[@class='ant-btn css-dts6b9 ant-btn-primary']//span[contains(text(),'確 定')]")


def coefficient_C4_upstreamEmissions_other(driver):
    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "validateOnly")))
    time.sleep(2)
    wait_and_click(driver, By.XPATH,
                   "//button[@class='ant-btn css-dts6b9 ant-btn-link d-flex align-items-center mt-2']")

    wait_and_click(driver, By.XPATH, "//div[@class='ant-collapse-item'][4]")

    wait_and_click(driver, By.XPATH, "//div[@class='ant-collapse-item']//span[contains(text(), '能資源')]")

    wait_and_click(driver, By.XPATH, "//a[@class='d-block my-2 text-dark'][3]")

    element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable(
        (By.XPATH, "//tr[@class='ant-table-row ant-table-row-level-0'][2]//span[@class='ant-checkbox-inner']")))
    actions = ActionChains(driver)
    actions.move_to_element(element).click().perform()
    time.sleep(1)

    wait_and_click(driver, By.XPATH,
                   "//button[@class='ant-btn css-dts6b9 ant-btn-primary']//span[contains(text(),'確 定')]")


def input_data_C3_C6(driver, source_type, rfg_or_fire):
    # random_letters = ''.join(random.choices(string.ascii_letters, k=10))

    validateOnlyGHGEvaluateItem = ["評估項目測試編輯"]
    validateOnlyMaterialNo = ["料號測試編輯"]
    validateOnlyMaterialName = ["物料名稱測試編輯"]
    validateOnlyCargoName = ["上游評估項目測試編輯"]
    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "validateOnly")))
    time.sleep(1)
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
        time.sleep(2)
    elif source_type == '營運產生之廢棄物 C4':
        coefficient_C4(driver)
        time.sleep(2)
    else:
        coefficient_C4_C6(driver)
        time.sleep(2)

    form_element = driver.find_element(By.ID, "validateOnly")  # 表單element

    input_elements = form_element.find_elements(By.TAG_NAME, "input")  # tag_name是input的

    new_inputelements = []

    for i in input_elements:  # 把disable的先移除掉
        if i.get_attribute("disabled") != None and i.get_attribute(
                "id") != "validateOnly_referenceFile" and i.get_attribute("type") != "search":
            print("")
        else:
            new_inputelements.append(i)

    time.sleep(2)
    for input_element in new_inputelements:
        try:
            if input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute(
                    "id") == 'validateOnly_GHGEvaluateItem':  # 評估項目
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                validateOnly_GHGEvaluateItem = source_type + validateOnlyGHGEvaluateItem[0]
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(validateOnly_GHGEvaluateItem)
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute(
                    "id") == 'validateOnly_MaterialNo':  # 料號
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(validateOnlyMaterialNo)
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute(
                    "id") == 'validateOnly_ActivityIntensityUnit':  # 活動強度(使用量)單位
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("Kg")
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute(
                    "id") == 'validateOnly_MaterialName':  # 物料名稱
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                validate_only_material_name = source_type + validateOnlyMaterialName[0]
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(validate_only_material_name)
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute(
                    "id") == 'validateOnly_CargoName':  # 上游評估項目
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(validateOnlyCargoName)
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute(
                    "id") == 'validateOnly_StaffOrDepartment':  # 上游評估項目
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("使用此方式之人原測試編輯")
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute(
                    "id") == 'validateOnly_JourneyNO':  # 上游評估項目
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("路程編號編輯測試")
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute(
                    "id") == 'validateOnly_StartLocation':  # 上游評估項目
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("運輸起點編輯測試")
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute(
                    "id") == 'validateOnly_EndLocation':  # 上游評估項目
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("運輸終點編輯測試")
                time.sleep(1)

            elif input_element.get_attribute("id") == 'validateOnly_ingredientName':
                continue

            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute(
                    "id") == 'validateOnly_ActivityIntensity':  # 活動強度(使用量)
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(50)
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute(
                    "id") == 'validateOnly_AnnualPurchaseAmount':  # 年度採購量
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(50)
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute(
                    "id") == 'validateOnly_TransportWeight':  # 上游重量
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(20)
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute(
                    "id") == 'validateOnly_TransportDistance':  # 上游運輸距離(km)
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(20)
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute(
                    "id") == 'validateOnly_TotalEmployees':  # 同行人數
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(20)
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute(
                    "id") == 'validateOnly_WorkingDays':  # 盤查期間次數
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(12)
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute(
                    "id") == 'validateOnly_AverageMovingDistance':  # 移動距離
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(50)
                time.sleep(1)

            elif input_element.get_attribute("type") == "search":  # 選單類 先click 再選 再click

                if input_element.get_attribute(
                        "aria-activedescendant") == "validateOnly_activityDataType_list_0":  # 活動數據種類
                    element = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '自動連續量測')]")))
                    element.click()
                    element = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '間歇量測')]")))
                    element.click()
                    time.sleep(1)
                elif input_element.get_attribute(
                        "aria-activedescendant") == "validateOnly_emitParaType_list_0":  # 排放係數種類
                    element = WebDriverWait(driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '自我/量測/質能')]")))
                    element.click()
                    element = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '同製程/設備經驗')]")))
                    element.click()
                    time.sleep(1)
                elif input_element.get_attribute("id") == "validateOnly_DataSource":
                    element = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '全球或區域級係數')]")))
                    element.click()
                    element = WebDriverWait(driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '國家級係數')]")))
                    element.click()
                    time.sleep(1)
                elif input_element.get_attribute("id") == "validateOnly_DataAttribute":
                    element = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '基於假設情境而來')]")))
                    element.click()
                    element = WebDriverWait(driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '可獲得特定廠址數據')]")))
                    element.click()
                    time.sleep(1)
                elif input_element.get_attribute("id") == "validateOnly_Commuting":  # 員工差旅交通方式
                    element = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '汽車-汽油')]")))
                    element.click()
                    element = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '計程車')]")))
                    element.click()
                    time.sleep(1)
                elif input_element.get_attribute("id") == "validateOnly_TransportTypeID":  # 運輸方式
                    element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                        (By.XPATH, "//span[contains(text(), '陸運,運輸 貨車 3.5-7.5噸 EURO5/RER S ')]")))
                    element.click()
                    element = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '陸運,無區分')]")))
                    element.click()
                    time.sleep(1)
                else:
                    continue
                time.sleep(1)

        except:  # StaleElementReferenceException
            print("")

    button = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']")))  # 送出
    button.click()
    time.sleep(3)
    if rfg_or_fire == 1:
        catch_response(driver, '人為逸散 C1(冷媒設備)')
    elif rfg_or_fire == 2:
        catch_response(driver, '人為逸散 C1(消防設備)')
    elif rfg_or_fire == 3:
        catch_response(driver, '輸入能源上游排放 C4(輸入電力上游)')
    elif rfg_or_fire == 4:
        catch_response(driver, '輸入能源上游排放 C4(其他輸入能源上游)')
    elif rfg_or_fire == 0:
        catch_response(driver, source_type)
    time.sleep(2)


def input_data_C4(driver, source_type, elec_other):
    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "validateOnly")))
    time.sleep(1)
    if elec_other == 1:
        coefficient_C4_upstreamEmissions_elec(driver)
    elif elec_other == 2:
        coefficient_C4_upstreamEmissions_other(driver)

    form_element = driver.find_element(By.ID, "validateOnly")  # 表單element

    input_elements = form_element.find_elements(By.TAG_NAME, "input")  # tag_name是input的

    new_inputelements = []

    for i in input_elements:  # 把disable的先移除掉
        if i.get_attribute("disabled") != None and i.get_attribute(
                "id") != "validateOnly_referenceFile" and i.get_attribute("type") != "search":
            print("")
        else:
            new_inputelements.append(i)

    time.sleep(2)
    for input_element in new_inputelements:
        try:
            if input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute(
                    "id") == 'validateOnly_MaterialNo':  # 料號
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("料號測試")
                time.sleep(1)
            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute(
                    "id") == 'validateOnly_ActivityIntensityUnit':  # 活動強度(使用量)單位
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("Kg")
                time.sleep(1)

            elif input_element.get_attribute("placeholder") == "請輸入" and input_element.get_attribute(
                    "id") == 'validateOnly_Scenario':  # 上游評估項目
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys("情境說明編輯測試")
                time.sleep(1)

            elif input_element.get_attribute("id") == 'validateOnly_ingredientName':
                continue

            elif input_element.get_attribute("placeholder") == "請輸入數字" and input_element.get_attribute(
                    "id") == 'validateOnly_ActivityIntensity':  # 活動強度(使用量)
                input_element.send_keys(Keys.CONTROL, 'a')
                input_element.send_keys(Keys.BACKSPACE)
                WebDriverWait(driver, 5).until(EC.visibility_of(input_element)).send_keys(500)
                time.sleep(1)


            elif input_element.get_attribute("type") == "search":  # 選單類 先click 再選 再click

                if input_element.get_attribute(
                        "aria-activedescendant") == "validateOnly_activityDataType_list_0":  # 活動數據種類
                    element = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '自動連續量測')]")))
                    element.click()
                    element = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '間歇量測')]")))
                    element.click()
                    time.sleep(1)
                elif input_element.get_attribute(
                        "aria-activedescendant") == "validateOnly_emitParaType_list_0":  # 排放係數種類
                    element = WebDriverWait(driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '自我/量測/質能')]")))
                    element.click()
                    element = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '同製程/設備經驗')]")))
                    element.click()
                    time.sleep(1)
                elif input_element.get_attribute("id") == "validateOnly_DataSource":
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '全球')]")))
                    element.click()

                elif input_element.get_attribute("id") == "validateOnly_DataAttribute":
                    input_element.click()
                    element = WebDriverWait(driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '基於')]")))
                    element.click()

                else:
                    continue
                time.sleep(1)

        except:  # StaleElementReferenceException
            print("")

    button = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']")))  # 送出
    button.click()
    time.sleep(3)
    if elec_other == 1:
        catch_response(driver, '輸入能源上游排放 C4(輸入電力上游)')
    elif elec_other == 2:
        catch_response(driver, '輸入能源上游排放 C4(其他輸入能源上游)')
    time.sleep(2)


def workhour_edit(driver, url):
    element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '編輯')]")))
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '編輯')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(2)
    input_elements = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.XPATH, '//input[@placeholder="請輸入"]'))
    )
    # 輸入不同的值到每個 <input> 元素
    for i, element in enumerate(input_elements):
        if i < 96:
            element.send_keys(Keys.CONTROL, 'a')  # crtl+A 全選
            element.send_keys(Keys.BACKSPACE)  # 刪除
            element.send_keys(24)
        else:
            value_to_input = "工時測試"
            element.send_keys(Keys.CONTROL, 'a')  # crtl+A 全選
            element.send_keys(Keys.BACKSPACE)  # 刪除
            element.send_keys(value_to_input)

    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '自動連續量測')]")))
    element.click()
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '間歇量測')]")))
    element.click()
    time.sleep(1)
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '自我/量測/質能')]")))
    element.click()
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '同製程/設備經驗')]")))
    element.click()
    element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '儲存')]")))
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '儲存')]")))
    driver.execute_script("arguments[0].click()", element)
    catch_response(driver, '人為逸散 C1(工時輸入)')  # 抓response


def elec_edit(driver, url):
    time.sleep(2)
    page_height = driver.execute_script(
        "return Math.max( document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight );")
    # 往下滾動至頁面中間
    driver.execute_script(f"window.scrollTo(0, {page_height // 3});")
    time.sleep(2)
    element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '編輯')]")))
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '編輯')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(2)
    input_elements = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.XPATH, '//input[@placeholder="請輸入"]'))
    )
    # 輸入不同的值到每個 <input> 元素
    for i, element in enumerate(input_elements):
        if i < 78:
            element.send_keys(Keys.CONTROL, 'a')  # crtl+A 全選
            element.send_keys(Keys.BACKSPACE)  # 刪除
            element.send_keys(200)
        else:
            value_to_input = "電力編輯測試"
            element.send_keys(Keys.CONTROL, 'a')  # crtl+A 全選
            element.send_keys(Keys.BACKSPACE)  # 刪除
            element.send_keys(value_to_input)
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '自動連續量測')]")))
    element.click()
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '間歇量測')]")))
    element.click()
    time.sleep(1)
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '自我/量測/質能')]")))
    element.click()
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '同製程/設備經驗')]")))
    element.click()
    element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '儲存')]")))
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '儲存')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(5)
    catch_response(driver, '輸入電力 C1')  # 抓response


def steam_edit(driver, url):
    time.sleep(1)
    edit_steam_input(driver, url)  # 做蒸氣加項
    element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, '//div[text()="蒸氣減項 B.3.2.b"]')))
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//div[text()="蒸氣減項 B.3.2.b"]')))
    driver.execute_script("arguments[0].click()", element)

    edit_steam_input(driver, url)  # 做蒸氣減項


def edit_steam_input(driver, url):
    element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '編輯')]")))
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '編輯')]")))
    driver.execute_script("arguments[0].click()", element)

    input_elements = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.XPATH, '//input[@placeholder="請輸入"]'))
    )
    # 輸入不同的值到每個 <input> 元素
    for i, element in enumerate(input_elements):
        if i < 12:
            element.send_keys(Keys.CONTROL, 'a')  # crtl+A 全選
            element.send_keys(Keys.BACKSPACE)  # 刪除
            element.send_keys(200)
        else:
            value_to_input = '蒸氣測試'
            element.send_keys(Keys.CONTROL, 'a')  # crtl+A 全選
            element.send_keys(Keys.BACKSPACE)  # 刪除
            element.send_keys(value_to_input)
    time.sleep(2)
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '自動連續量測')]")))
    element.click()
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '間歇量測')]")))
    element.click()
    time.sleep(1)
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '自我/量測/質能')]")))
    element.click()
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '同製程/設備經驗')]")))
    element.click()
    element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, "//span[contains(text(), '儲存')]")))
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '儲存')]")))
    driver.execute_script("arguments[0].click()", element)
    time.sleep(5)
    catch_response(driver, '輸入蒸汽')  # 抓response