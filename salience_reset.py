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

def salience_reset(driver,url):        
    url = url + "project/risk/806"
    has_next_page = True
    while has_next_page:
        
        reset_buttons = driver.find_elements(By.XPATH, '//button[contains(text(), "清空評分紀錄")]')
        for button in reset_buttons: 
            driver.execute_script("arguments[0].click();", button)
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '確 定')]"))).click()
            time.sleep(1)
       

            
        try:
            next_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//li[@title='下一頁']//button[@class='ant-pagination-item-link']")))
            ActionChains(driver).move_to_element(next_button).click().perform()
            
            time.sleep(2)
            
        except:
            has_next_page = False