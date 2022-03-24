import datetime
import logging
import os
import re
import sys
import time
import traceback
from datetime import datetime
import re
import pandas as pd
import win32com.client as w3c
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait as wait
from scipy.io import wavfile
import io
import requests
import pickle
import numpy as np
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import StaleElementReferenceException

class RPAInputVal:
    def __init__(self, mapping_book_loc='.'):
        self.l_dict = pickle.load(open(os.path.join(mapping_book_loc,'l_dict'),'rb'))
        self.v_dict = pickle.load(open(os.path.join(mapping_book_loc,'v_dict'),'rb'))
        self._create_arr()
    def _create_arr(self):
        r = len(self.l_dict)
        self.sp = min([v for v in self.l_dict.values()])
        self.arr = np.zeros((r, self.sp))
        self.mapping_lst = []
        for ind, c in enumerate(self.v_dict):
            self.arr[ind] = self.v_dict[c][:self.sp]
            self.mapping_lst.append(c)
    def run(self, wav, num_of_str=5):
        result = ''
        for _ in range(num_of_str):
            c = self.mapping_lst[np.argmin(abs(wav[:self.sp] - self.arr).sum(axis=1))]
            wav = wav[self.l_dict[c]:]
            result += c
        return result

class address_getter:
    def __init__(self,driver,city,region,road):
        self.driver = driver
        self.city = city
        self.region = region
        self.road = road
        self.total_result_df = pd.DataFrame()

    def address_typer(self, driver, city, region, road):
        #不知為啥只能怎樣 反正要把字串符號加進去
        city = f"""\"{city}\""""
        driver.find_element(By.XPATH, (f"""//area[contains(@title, {city})]""")).click()
        region = f"""{region}"""
        
        #選擇區域(下拉式選單)
        wait(self.driver, 10).until(EC.presence_of_element_located((By.XPATH, f"""//select[@id='areaCode']""")))
        region_select = Select(driver.find_element(By.XPATH, (f"""//select[@id='areaCode']""")))
        region_select.select_by_visible_text(f"""{region}""")
        #寫路
        driver.find_element(By.XPATH, (f"""//input[@name='street']""")).send_keys(road)
        
    def get_current_captcha(self, driver):
        # init 一次就可以
        tool = RPAInputVal(mapping_book_loc='.')
        browser_cookies ={i['name']: i['value'] for i in driver.get_cookies()}
        # 以下辨識驗證碼
        # new_cap = driver.find_element_by_xpath('//*[@id="imageBlock_captchaKey"]')
        cap_key = driver.find_element(By.XPATH, (f"""//*[@id="captchaKey_captchaKey"]"""))
        voice_src = 'https://www.ris.gov.tw/info-doorplate/captcha/sound/{}/{}'.format(cap_key.get_attribute('value'),int(time.time()*1000))
#         print(voice_src)
        voice_res = requests.get(voice_src, cookies=browser_cookies)
        result = tool.run(wavfile.read(io.BytesIO(voice_res.content))[1])

        return result   
    
    def run(self):
        
        self.address_typer(self.driver, self.city, self.region, self.road)
        
        while True:
            try:
                
                #取得captcha
                #反覆測試機制 不然有時候會出錯
                while True:
                    try:
                        captcha = self.get_current_captcha(self.driver)
                    except:
                        continue
                    break
                
                self.driver.find_element(By.XPATH, (f"""//input[@id="captchaInput_captchaKey"]""")).send_keys(captcha)
                self.driver.find_element(By.XPATH, (f"""//button[@id="goSearch"]""")).click()
                wait(self.driver, 15).until(EC.presence_of_element_located((By.XPATH, f"""//td[@data-jqlabel='門牌資料']""")))
            except (TimeoutException, ConnectionError) as e:
                # 如果失敗的話(可能偶爾認證碼錯誤之類的吧)就重試
                self.driver.find_element(By.XPATH, (f"""//input[@id="captchaInput_captchaKey"]""")).clear()
                
                # 查無資料的情況
                try:
                    self.driver.find_element(By.XPATH, (f"""//button[@class="swal2-confirm swal2-styled"]""")).click()
                    # 沒有查無資料確定的按鈕就跳出本次迴圈，繼續搜下一個地區
                    return pd.DataFrame()
                except NoSuchElementException as e:
                    continue
                    
                continue
            break
            
        # 抓取所有結果，存入address_list
        address_list = []
        # 每五十個結果就要按下一頁，直到不能為止
        # while wait(driver,5).until(EC.element_to_be_clickable((By.XPATH, "//td[@id='next_result-pager']"))):
        while True:
            # 拿回所有門牌的結果
            result_address = self.driver.find_elements(By.XPATH, (f"""//td[@data-jqlabel='門牌資料']"""))

            try:
                for i in result_address:
                    address_list.append(i.text)
            except StaleElementReferenceException as e: 
                # 可能有時候載太慢了會出現這個錯誤 就重新把地址在裝一次吧 反正有重疊之後再drop duplicate
                continue                

            #拿到結果，按下一頁(超過五十個結果要額外按下一頁)
            try:
                self.driver.find_element(By.XPATH, f"""//td[@id='next_result-pager']""").click()
            except:
                #如果不能按下一頁(即最後一頁)則停止蒐資料
                break
            break
            
        #做dict以方便存成df
        result_dict = {
            "city" : [city] * len(address_list),
            "address" : address_list
        }
        
        return pd.DataFrame.from_dict(result_dict)

if __name__ == "__main__":
    # 全台灣區域路檔案
    path_df = pd.read_csv('opendata110road.csv')
    
    # 0只是欄首的中文翻譯,drop掉
    path_df.drop(index=path_df.index[0], 
            axis=0, 
            inplace=True)
    path_df.reset_index(drop=True)
    
    # 取區域，即site_id為後三碼
    path_df['region'] = path_df['site_id'].str[-3:]
    
    current_version = "99"
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    # chrome_options.add_argument('--start-maximized')
    chromedriver_path = f"\\\\eip.esunbank.com.tw@SSL\DavWWWRoot\sites\C010\DocLib1\kentsai\chromedriver\{current_version}\chromedriver.exe"
    # chromedriver_path = f"""chromedriver.exe"""
    
    # init index
    idx = 0
    
    # 輸入之前做到的index
    # 看資料夾下面的pickle的檔名數字
    start_idx = 0
    print(start_idx)
    
    driver = webdriver.Chrome(chromedriver_path)
    
    #生成住址檔案
    total_result_df = pd.DataFrame()
    
    for idx in tq.tqdm(range(start_idx-1, len(path_df))):
    # for idx in tq.tqdm(range(5)):
        while True:
        #     #每一百次重啟一個，感覺重開比較沒什麼問題
        #     if(idx%100==0):
        #         driver.quit()
        #         driver = webdriver.Chrome(chromedriver_path)
            
            try:
                # 開啟門牌第一頁頁面
                driver.get("https://www.ris.gov.tw/info-doorplate/app/doorplate/main")
                #點擊門牌查詢
                driver.find_element(By.CSS_SELECTOR, "[data-type='doorplate']").click()

                #開始點地圖
                city = path_df['city'].iloc[idx]
                region = path_df['region'].iloc[idx]
                road = path_df['road'].iloc[idx]

                address_getter_tool = address_getter(driver,city,region,road)
                
                total_result_df = total_result_df.append(address_getter_tool.run())
                
            except Exception as e:
                # 有時候還沒載入完畢或是502 bad gateway就重新執行吧 或其他各種疑難雜症
                continue
            break
            
    total_result_df.to_pickle(f"""addres_idx_{start_idx}_to_{idx}.p""")
    
    
    