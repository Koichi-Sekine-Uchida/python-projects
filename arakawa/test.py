import win32com.client as win32
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.alert import Alert
from webdriver_manager.microsoft import EdgeChromiumDriverManager



from selenium import webdriver
import pandas as pd
import time

#ドライバーの自動インストール
driver = webdriver.Edge(EdgeChromiumDriverManager().install())


# Excelファイルを読み込む
df = pd.read_excel('C:\tools\selenium\arakawa\filtered_data.xlsx')

# WebDriverの設定
driver = webdriver.Chrome('/path/to/chromedriver')

# 各グループIDに対してループ処理
for group_id in df['グループID']:
    # ウェブサイトにアクセス
    driver.get('https://school.edumall.jp/dlvr/CAFS15003')

    # グループIDを入力し、検索ボタンをクリック
    driver.find_element_by_id('group_id_input_field').send_keys(group_id)
    driver.find_element_by_id('search_button').click()
    
    # 結果の読み込みに少し待機
    time.sleep(2)

    # 検索結果を解析して操作を行う（具体的な要素のIDやクラス名は実際のサイトに合わせて変更する必要がある）
    # 例: 配信済の行があるかどうかチェックし、あればその配信IDのリンクをクリック
    # ...

# ブラウザを閉じる
driver.quit()
