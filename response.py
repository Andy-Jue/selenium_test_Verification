from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import logging

def catch_response(driver,source_type):
    try:
        time.sleep(2)
        message_element = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CLASS_NAME, 'ant-notification-notice-message')))
        # 獲取 message 內容
        message = message_element.text
        print(message)
        # 輸出 message 內容
        message=message+'(' + source_type + ')'
        print(f"Message 內容：{message}")
        if '成功' in message:
            logging.info(message)
        else:
            logging.error(message)
    except:
        print("無提示訊息")

