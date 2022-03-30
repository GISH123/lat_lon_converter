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
    
    # 在資料夾抓門牌住址的pickle檔
    files = []
    for file in os.listdir("./address"):
        if file.endswith(".p"):
            files.append(file)
    
    rootdir = r'./address/'

    #抓取所有門牌地址的pickle檔
    for filename in files:
        
        lat_list = []
        lon_list = []
        total_result_df = pd.read_pickle(rootdir + filename).reset_index(drop=True)
        
        #每一個pickle檔的門牌地址，抓經緯度
        for i in range(len(total_result_df)):
            
            # 檔案儲存的起始index
            start_i = 0
            
            while True:

                #每一百次重啟一個，因為我不知道為什麼會怪怪的，地圖會開始卡在定位中
                if(i%100==0):
                    lat_lon_driver.quit()
                    lat_lon_driver = webdriver.Chrome(chromedriver_path)
                    print(i)

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

                    #re.findall吐出來是list 但只會抓到一個數字，就先這樣抓
                    lat_list.append(lat[0])
                    lon_list.append(lon[0])
                    
                except TimeoutException:
                    #有問題就換下一個
                    lat_list.append(None)
                    lon_list.append(None)                
                    break
                break
                
            # 5000筆存一次
            if((i%5000==0) or (i == len(total_result_df)-1) ):
                # loc的i會包含末位..
                total_result_df.loc[start_i:i,'lat'] = lat_list
                total_result_df.loc[start_i:i,'lon'] = lon_list

                #儲存pickle
                total_result_df[start_i:i+1].to_pickle("./processed/" + filename + \
                                                       '_lat_lon_processed_' + str(start_i) +'_to_' + str(i) +'.pickle')
                
                #更新下一次存檔的start_i
                start_i = i
        
            
            