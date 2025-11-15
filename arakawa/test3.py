from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
import time

# Excelファイルの読み込み
df = pd.read_excel("C:\tools\selenium\arakawa\test.xlsx", usecols="A") # ExcelファイルのパスとA列を指定

# Seleniumの設定
driver = webdriver.Chrome(executable_path="C:\Program Files\PackageManagement\NuGet\Packages\Selenium.WebDriver.ChromeDriver.116.0.5845.9600") # chromedriverのパスを指定
driver.get("https://school.edumall.jp/dlvr/CAFS15003") # 操作したいWebページのURLを指定

# Excelのデータを1行ずつ処理
for index, row in df.iterrows():
    # Web上の欄にデータを設定
    input_element = driver.find_element(By.ID, "input_element_id") # 入力欄のIDを指定
    input_element.clear()
    input_element.send_keys(str(row[0]))
    
    # 検索ボタンをクリック
    search_button = driver.find_element(By.ID, "search_button_id") # 検索ボタンのIDを指定
    search_button.click()
    
    time.sleep(2) # 必要に応じて待機時間を設定
    
    # 再配信チェックをチェック
    redelivery_checkbox = driver.find_element(By.ID, "redelivery_checkbox_id") # 再配信チェックボックスのIDを指定
    redelivery_checkbox.click()
    
    # 更新ボタンをクリック
    update_button = driver.find_element(By.ID, "update_button_id") # 更新ボタンのIDを指定
    update_button.click()
    
    time.sleep(2) # 必要に応じて待機時間を設定

# 終了処理
driver.quit()
