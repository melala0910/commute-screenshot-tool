# -*- coding: utf-8 -*-
"""
Created on Wed Mar 12 12:06:17 2025

@author: zijie.liu
"""
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import urllib.parse
import time 
import pandas as pd
import os
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
import re

file_path = r"C:\Users\zijie.liu\Desktop\selenium\公里數.xlsx"

df = pd.read_excel(file_path)
try:
    # 取得檔案所在的資料夾路徑
    folder_path = os.path.dirname(file_path)

    # 創建 "截圖" 資料夾
    screenshot_folder = os.path.join(folder_path, "公里數截圖")
    os.makedirs(screenshot_folder, exist_ok=True)  # exist_ok=True 表示如果資料夾已存在，不會引發錯誤

    print(f"已在 {folder_path} 中創建 '公里數截圖' 資料夾。")

except FileNotFoundError:
    print(f"找不到檔案：{file_path}")
except Exception as e:
    print(f"讀取檔案或創建資料夾時發生錯誤：{e}")

# 新增三個欄位，初始化為 None
df['公里1'] = None
df['公里2'] = None
df['公里3'] = None

origin_address = pd.read_excel(r"C:\Users\zijie.liu\Desktop\selenium\原始地址.xlsx")
origin_address = origin_address['地址']
# 使用 lambda 函數和正規表達式直接應用到 df.iloc[:, 2]
df.iloc[:, 2] = df.iloc[:, 2].apply(lambda address: re.sub(r'\d+鄰', '', address))



# 起點和終點地址
start_address = "台中市北屯區崇德六路一段59號10"
end_address = "421台中市后里區后科南路28號"

# URL 編碼
encoded_start = urllib.parse.quote(start_address)
encoded_end = urllib.parse.quote(end_address)

# 構建 Google 地圖 URL
maps_url = f"https://www.google.com.tw/maps/dir/{encoded_start}/{encoded_end}"

# 初始化 WebDriver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)


# 開啟 Google 地圖
driver.get(maps_url)

start_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="sb_ifc50"]/input')))
Target_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="sb_ifc51"]/input')))

#交通方式
driving = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="omnibox-directions"]/div/div[2]/div/div/div/div[2]/button/div[1]/span[3]')))
motorcycle = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="omnibox-directions"]/div/div[2]/div/div/div/div[3]/button/div[1]/span[3]')))
train = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="omnibox-directions"]/div/div[2]/div/div/div/div[4]/button/div[1]/span[3]')))
walk = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="omnibox-directions"]/div/div[2]/div/div/div/div[5]/button/div[1]')))
bike = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="omnibox-directions"]/div/div[2]/div/div/div/div[6]/button/div[1]/span[3]')))


# 使用 JavaScript 調整縮放比例
driver.execute_script("document.body.style.zoom='80%'")
driver.maximize_window()

# 清除輸入框內容
start_input.clear()
Target_input.clear()

'----------------------以上為基礎設置---------------------------'
start_time = time.time() # 執行時間
new_data = []
for i in range(len(df)):
    df_id = df.iloc[i,0] # id
    df_way = df.iloc[i,1] #交通方式
    df_start = df.iloc[i,2] # 起點
    df_target = df.iloc[i,3] # 終點
    df_km1 = df.iloc[i,4] #公里1
    df_km2 = df.iloc[i,5] #公里2
    df_km3 = df.iloc[i,6] #公里3
    df_address = origin_address[i]
    '--------開始搜尋地址------'
    
    print('---------------------------------------------')
    print('正在搜尋 :',df_id,df_way,'起始位置 :',df_start)
    time.sleep(1)
    start_input.clear()
    time.sleep(1)
    Target_input.clear()
    
    time.sleep(3)
    start_input.send_keys(df_start)
    time.sleep(0.5)
    Target_input.send_keys(df_target, Keys.ENTER)
    time.sleep(3)
    '--------抓公里數---------'
    km_1 = None
    km_2 = None
    km_3 = None
    
    
    if df_way == "開車" or df_way == "汽車":
        driving.click()
        time.sleep(1.5)  
    elif df_way == "機車":
        motorcycle.click()
        time.sleep(1.5)
    elif df_way == "火車":
        train.click()
        time.sleep(1.5)
    elif df_way == "走路":
        walk.click()
        time.sleep(1.5)   
            
    # 公里1 -------------------
    try:
        # 嘗試第一組 XPath
        km_1 = WebDriverWait(driver,5).until(EC.presence_of_element_located((By.XPATH, "//*[@id='section-directions-trip-0']/div[1]/div/div[1]/div[2]/div")))
        print("找到第一組 XPath ：", km_1.text)
        df_km1 = km_1.text
        
    except TimeoutException:
        try:
            # 嘗試第二組 XPath
            km_1 = WebDriverWait(driver,1).until(EC.presence_of_element_located((By.XPATH, "//*[@id='section-directions-trip-0']/div[1]/div/div[1]/div[2]")))
            print("找到第二組 XPath ：", km_1.text)
            df_km1 = km_1.text
            
        except TimeoutException:
            print('沒有找到xpath')
            pass
    except Exception as e:
        print(f"發生錯誤：{e}")
          
    
    # 公里2 --------------------
    try:
        # 嘗試第一組 XPath
        km_2 = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH,'//*[@id="section-directions-trip-1"]/div[1]/div/div[1]/div[2]/div' )))
        print("找到第一組 XPath ：", km_2.text)
        df_km2 = km_2.text
        
    except TimeoutException:
        try:
            # 嘗試第二組 XPath
            km_2 = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH,'//*[@id="section-directions-trip-1"]/div[1]/div/div[1]/div[2]')))
            print("找到第二組 XPath：", km_2.text)
            df_km2 = km_2.text
            
        except TimeoutException:
            print('沒有找到xpath')
            pass
    except Exception as e:
        print(f"發生錯誤：{e}")
    
    
    # 公里3 -------------------
    try:
        # 嘗試第一組 XPath
        km_3 = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH,'//*[@id="section-directions-trip-2"]/div[1]/div/div[1]/div[2]/div' )))
        print("找到第一組 XPath：", km_3.text)
        df_km3 = km_3.text
        
    except TimeoutException:
        try:
            # 嘗試第二組 XPath
            km_3 = WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH,'//*[@id="section-directions-trip-2"]/div[1]/div/div[1]/div[2]')))
            print("找到第二組 XPath：", km_3.text)
            df_km3 = km_3.text
            
        except TimeoutException:
            print('沒有找到xpath')
            pass
    except Exception as e:
        print(f"發生錯誤：{e}")
        
    # 截圖--------------------------
    
    try:
        # 創建資料夾（如果不存在）
        if not os.path.exists(screenshot_folder):
            os.makedirs(screenshot_folder)

        # 構建檔案名稱
        file_name = str(df_id) + '_' + str(df_way) + '_' + str(df_address) + '.png'
        # 構建完整的檔案路徑
        screenshot_path = os.path.join(screenshot_folder, file_name)

        # 保存截圖
        driver.save_screenshot(screenshot_path)

        print(f"截圖已保存到：{screenshot_path}")

    except Exception as e:
        print(f"截圖時發生錯誤：{e}")
    
    new_data.append([df_id, df_way, df_start, df_target,df_km1,df_km2,df_km3])

# 將 new_data list 轉換為新的 DataFrame
new_df = pd.DataFrame(new_data, columns=['index', '交通方式', '起點', '終點','公里1','公里2','公里3'])

end_time = time.time()
execution_time = int(round(end_time - start_time,0))



# 計算執行時間

print('-------------------------------------------------------------')
def convert_seconds(seconds):
    minutes = int(round(seconds // 60,0))  # 計算分鐘數
    remaining_seconds = seconds % 60  # 計算剩餘的秒數
    return f"搜尋執行時間 : {minutes}分 {remaining_seconds}秒"
print(convert_seconds(execution_time))


file_name="公里數_result.csv"
new_df_path = os.path.join(folder_path, file_name)
new_df.to_csv(new_df_path, encoding='utf-8-sig', index=False)

print('檔案已保存至: ',new_df_path)




 
