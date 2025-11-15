import win32com.client as win32
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Edgeドライバーのインストール
driver_path = EdgeChromiumDriverManager().install()
service = Service(driver_path)
driver = webdriver.Edge(service=service)

# 設定ファイルのパス
filename = r"C:\tools\python-projects\EduMall\ACCIS注文連携データ取得\配信設定ファイル.xlsx"

# ファイル確認
if not os.path.exists(filename):
    print(f"エラー: 指定されたExcelファイルが見つかりません: {filename}")
    exit(1)

# Excelを開く
xlApps = win32.Dispatch("Excel.Application")
workbook = xlApps.Workbooks.Open(filename)
sheet = workbook.Worksheets("Sheet1")

# Excelデータ取得
sellSide_url = str(sheet.Cells(1, 2).Value)  # アクセスするリンク
edumall_id = str(sheet.Cells(2, 2).Value)  # EduMallのID
edumall_pw = str(sheet.Cells(3, 2).Value)  # EduMallのPW
sleep_time = int(sheet.Cells(4, 2).Value)
school_name = str(sheet.Cells(5, 2).Value)  # 学校名
title = str(sheet.Cells(6, 2).Value)  # タイトル
end_year = str(sheet.Cells(7, 2).Value)  # 設定する利用終了日
start_year = str(sheet.Cells(8, 2).Value)  # 利用開始日（年）
start_m = str(sheet.Cells(9, 2).Value)  # 利用開始日（月）
start_d = str(sheet.Cells(10, 2).Value)  # 利用開始日（日）

# アクセスするリンク
driver.get(sellSide_url)
time.sleep(3)

# ログイン処理
driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[1]/input").send_keys(edumall_id + Keys.TAB)
driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[2]/input").send_keys(edumall_pw)
driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[3]/button").click()

# **📌 メニュー (`CAdMenu.jsp`) の `iframe` に切り替え**
try:
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "menu")))
    driver.switch_to.frame("menu")
except:
    print("Error: menu iframe not found.")
    exit(1)

# 6. 「注文管理」をクリック
try:
    order_menu = driver.find_element(By.XPATH, '//p[@onclick="openMenu(\'3\')"]')
    driver.execute_script("arguments[0].click();", order_menu)
    time.sleep(1)
except:
    print("Error: 注文管理メニュー not found.")
    driver.quit()
    exit(1)

# 7. 「ACCIS注文連携」をクリック
try:
    accis_menu = driver.find_element(
        By.XPATH,
        '//a[@onclick="showPage(this, \'order/COdAccisOrderMatch.jsp\'); return false;"]'
    )
    driver.execute_script("arguments[0].click();", accis_menu)
    time.sleep(1)
except:
    print("Error: ACCIS注文連携 not found.")
    driver.quit()
    exit(1)

# **📌 `center` の `iframe` に切り替え**
driver.switch_to.default_content()
try:
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "center")))
except:
    print("Error: center iframe not found.")
    exit(1)

'''
# **📌 検索フォームにEXCELのデータを入力**
try:
    driver.find_element(By.XPATH, "/html/body/div/form/table[1]/tbody/tr[1]/td[1]/input").send_keys(school_name)
    driver.find_element(By.XPATH, "/html/body/div/form/table[1]/tbody/tr[2]/td[2]/input").send_keys(title)
    driver.find_element(By.XPATH, "/html/body/div/form/table[1]/tbody/tr[4]/td/input[1]").send_keys(start_year)
    driver.find_element(By.XPATH, "/html/body/div/form/table[1]/tbody/tr[4]/td/input[2]").send_keys(start_m)
    driver.find_element(By.XPATH, "/html/body/div/form/table[1]/tbody/tr[4]/td/input[3]").send_keys(start_d)
    print("検索条件入力完了.")
except:
    print("Error: フォーム入力に失敗しました.")
    exit(1)
'''

# **📌 検索ボタンを押下**
try:
    search_button = driver.find_element(By.XPATH, "/html/body/div/form/table[2]/tbody/tr/td/input[1]")
    search_button.click()
    time.sleep(3)  # 検索結果が表示されるのを待つ
except:
    print("Error: 検索ボタンが見つかりません。")
    exit(1)

# ── 検索結果表示後 （113行目以降） ──

# ── for page 以降の全コード ──

import pandas as pd
import os

COLUMN_NAMES = [
    '受注番号','学校名(ウチダ学校コード)','タイトル','学校名','タイトル',
    '受注明細番号','処理区分','型番(数量)','利用期間','型番(数量)','利用期間'
]
records = []

# center iframe に切り替え
driver.switch_to.default_content()
driver.switch_to.frame("center")

# ページャの a 要素を一度だけ取得（1～16ページなら 16 要素のはず）
pager_links = WebDriverWait(driver, 10).until(
    EC.presence_of_all_elements_located(
        (By.XPATH, "//div[contains(@class,'paging')]//a")
    )
)

# 取得した要素を順番にクリックしてテーブルをスクレイピング
for idx, link in enumerate(pager_links, start=1):
    # スクロールしてビュー内に入れる
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", link)
    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, f"(//div[contains(@class,'paging')]//a)[{idx}]")))
    # クリック
    link.click()
    time.sleep(1)

    # テーブル本体を取得（実際の class/XPath に合わせてください）
    table = driver.find_element(By.XPATH, "//table[contains(@class,'list-data')]")
    rows = table.find_elements(By.TAG_NAME, "tr")[1:]

    for row in rows:
        cols = row.find_elements(By.TAG_NAME, "td")
        if len(cols) < len(COLUMN_NAMES):
            continue
        row_data = { COLUMN_NAMES[i]: cols[i].text.strip() for i in range(len(COLUMN_NAMES)) }
        records.append(row_data)

# ブラウザを閉じる
driver.quit()

# DataFrame→Excel
df = pd.DataFrame(records, columns=COLUMN_NAMES)
out = os.path.join(os.getcwd(), "accis_orders.xlsx")
df.to_excel(out, index=False)
print(f"全{len(records)}件を '{out}' に出力しました。")
