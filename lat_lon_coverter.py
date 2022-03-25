import datetime
import logging
import os
import re
import sys
import time
import traceback
from datetime import datetime

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
import re
from selenium.common.exceptions import TimeoutException

if __name__ == "__main__":
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    # chrome_options.add_argument('--start-maximized')
    # chromedriver_path = f"\\\\eip.esunbank.com.tw@SSL\DavWWWRoot\sites\C010\DocLib1\kentsai\chromedriver\{current_version}\chromedriver.exe"
    chromedriver_path = f"""chromedriver.exe"""
    
    lat_lon_driver = webdriver.Chrome(chromedriver_path,chrome_options=chrome_options)
    
    # 讀入門牌住址檔案
    total_result_df = pd.read_pickle("addres_idx_0_to_65.p")
    
    # 生成經緯度
    lat_list = []
    lon_list = []
    
    idx = 0
    
    # for i in range(len(total_result_df)):
    for i in range(idx,len(total_result_df)):
        while True:
            
            #每一百次重啟一個，因為我不知道為什麼會怪怪的，地圖會開始卡在定位中
            if(i%100==0):
                lat_lon_driver.quit()
                lat_lon_driver = webdriver.Chrome(chromedriver_path)
                
            try:
                lat_addr = total_result_df['address'].iloc[i]
                lat_lon_driver.get(f"""https://map.tgos.tw/TGOSCloud/Web/Map/TGOSViewer_Map.aspx?addr={lat_addr}""")


                lat_lon_text_div = wait(lat_lon_driver, 15).until(
                    # 找到有lat lon的div
                    EC.presence_of_element_located((By.XPATH, f"""//*[@id="MapBox"]/div[1]/div[2]/div"""))
                )

                lat_lon_text_div_text = lat_lon_text_div.text.split("\n")
                lat_text = lat_lon_text_div_text[1]
                lon_text = lat_lon_text_div_text[2]

                lat = re.findall(r"[-+]?(?:\d*\.\d+|\d+)", lat_text)
                lon = re.findall(r"[-+]?(?:\d*\.\d+|\d+)", lon_text)    

                lat_list.append(lat)
                lon_list.append(lon)
                if(i%10==0):
                    print(i)
            except TimeoutException:
                #有問題就換下一個
                break
            break
            
                
        total_result_df['lat'] = lat_list
        total_result_df['lon'] = lon_list
    
    total_result_df.to_pickle(f"""total_result_df_lat_lon{i}""")
    
        
        