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
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import WebDriverException
import sqlite3 as db
from sqlite3 import IntegrityError

if __name__ == "__main__":
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    # chrome_options.add_argument('--start-maximized')
    # chromedriver_path = f"\\\\eip.esunbank.com.tw@SSL\DavWWWRoot\sites\C010\DocLib1\kentsai\chromedriver\{current_version}\chromedriver.exe"
    chromedriver_path = f"""chromedriver.exe"""
    
    # 在資料夾抓門牌住址的pickle檔

    # 先排序 確定順序一致
    all_address = pd.read_pickle('all_address.p').sort_values(by="address").reset_index(drop=True)

    #connect to database or create if doesn't exist
    conn = db.connect('lat_lon.db')
    
    #create cursor
    c = conn.cursor()

    #create table called 'lat_lon'
    c.execute("CREATE TABLE IF NOT EXISTS lat_lon (city TEXT, address TEXT PRIMARY KEY, lat REAL, lon REAL)")

    # 得到目前做到的部分
    count = pd.read_sql_query ('select count(*) from lat_lon', conn)
    start_index = int(count['count(*)'].loc[0])
    start_index
    
    lat_lon_driver = webdriver.Chrome(chromedriver_path,chrome_options=chrome_options)


    timeout_retry = 0
    row_list = []

    for i in range(start_index, len(all_address)):
    # for i in range(3):

    #     # 檔案儲存的起始index
    #     start_i = start_index

        while True:
            try:
                
                #每50次重啟一個，並commit，有時候跑太多次 網頁會卡住
                if(i%50==0):
                    c.executemany("INSERT INTO lat_lon VALUES (?, ?, ?, ?) ON CONFLICT(address) \
                    DO UPDATE SET city = 'inserted'", row_list)
                    # clear row list
                    row_list = []
                    conn.commit()
                    
    #                 conn.close()
    #                 #connect to database or create if doesn't exist
    #                 conn = db.connect('lat_lon.db')
    #                 #create cursor
    #                 c = conn.cursor()

                    lat_lon_driver.quit()
                    lat_lon_driver = webdriver.Chrome(chromedriver_path)
                    print(i)

                lat_addr = all_address['address'].iloc[i]
                lat_lon_driver.get(f"""https://map.tgos.tw/TGOSCloud/Web/Map/TGOSViewer_Map.aspx?addr={lat_addr}""")
                city = all_address['city'].iloc[i]

                lat_lon_text_div = wait(lat_lon_driver, 10).until(
                    # 找到有lat lon的div
                    EC.presence_of_element_located((By.XPATH, f"""//*[@id="MapBox"]/div[1]/div[2]/div"""))
                )

                lat_lon_text_div_text = lat_lon_text_div.text.split("\n")
                lat_text = lat_lon_text_div_text[1]
                lon_text = lat_lon_text_div_text[2]

                lat = re.findall(r"[-+]?(?:\d*\.\d+|\d+)", lat_text)
                lon = re.findall(r"[-+]?(?:\d*\.\d+|\d+)", lon_text)

                lat_val = lat[0]
                lon_val = lon[0]

                #re.findall吐出來是list 但只會抓到一個數字，就先這樣抓
                row_list.append((city, lat_addr, lat_val, lon_val))

                # c.execute(f"""INSERT INTO lat_lon VALUES ("{city}", "{lat_addr}", {lat_val}, {lon_val})""") # 不知為啥一筆一筆insert好像這樣做會有thread記憶體問題???

            except TimeoutException as t_e:
                #有問題就試三次，不行就換換下一個
                if(timeout_retry ==3):
                    timeout_retry = 0 
                    break
                timeout_retry += 1
                continue
            except WebDriverException as w_e:
                #有問題就試三次，不行就換換下一個
                if(timeout_retry ==3):
                    timeout_retry = 0 
                    break
                timeout_retry += 1
                continue
                
            # 改新寫法sqlite的upsert(on conflict)後，應該不會有這exception
            except IntegrityError as i_e:
                print(i)
                print("重複資料塞入，換另一個地址")
                break
            except KeyboardInterrupt:
                print('KeyboardInterrupt')
                sys.exit(0)
            except Exception as e: # 各式奇怪的疑難雜症
                print(e)
                #有問題就試三次，不行就換換下一個
                if(timeout_retry ==3):
                    timeout_retry = 0 
                    break
                timeout_retry += 1
                continue
            break

    #     # 5000筆存一次
    #     if((i%5000==0) or (i == len(total_result_df)-1) ):
    #         # loc的i會包含末位..
    #         total_result_df.loc[start_i:i,'lat'] = lat_list
    #         total_result_df.loc[start_i:i,'lon'] = lon_list

    #         #儲存pickle
    #         total_result_df[start_i:i+1].to_pickle("./processed/" + filename + \
    #                                                '_lat_lon_processed_' + str(start_i) +'_to_' + str(i) +'.pickle')
