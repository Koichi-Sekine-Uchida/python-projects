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
filename = r"C:\tools\python-projects\SELLSIDE\配信設定ファイル.xlsx"

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

# **📌 検索ボタンを押下**
try:
    search_button = driver.find_element(By.XPATH, "/html/body/div/form/table[2]/tbody/tr/td/input[1]")
    search_button.click()
    time.sleep(3)  # 検索結果が表示されるのを待つ
except:
    print("Error: 検索ボタンが見つかりません。")
    exit(1)

# **📌 表示された行数を取得**
try:
    rows = driver.find_elements(By.XPATH, "/html/body/form[2]/table[1]/tbody/tr")
    total_rows = len(rows) - 1  # ヘッダー行を除く
    print(f"検索結果の総件数: {total_rows} 件")
except:
    print("Error: 検索結果の行数を取得できませんでした。")
    exit(1)

# **📌 処理する行数をユーザーに入力させる**
while True:
    try:
        process_rows = int(input(f"処理する行数を入力してください (1～{total_rows}): "))
        if 1 <= process_rows <= total_rows:
            break
        else:
            print("範囲内の数値を入力してください。")
    except ValueError:
        print("数値を入力してください。")

# **📌 選択された行数だけ処理**
loopcounter = 5  # データ開始行
iframe = driver.find_element(By.XPATH, "/html/body/div/form/iframe")
driver.switch_to.frame(iframe)

for _ in range(process_rows):
    try:
        driver.find_element(By.XPATH, f"/html/body/form[2]/table[1]/tbody/tr[{loopcounter}]/td[4]/input[2]").clear()
        driver.find_element(By.XPATH, f"/html/body/form[2]/table[1]/tbody/tr[{loopcounter}]/td[4]/input[2]").send_keys(end_year)
        loopcounter += 2
    except:
        print(f"行 {loopcounter} の処理に失敗しました。")
        break

# **📌 送信処理**
try:
    down = driver.find_element(By.XPATH, "/html/body/form[2]/table[2]/tbody/tr/td[1]/input")
    driver.execute_script("arguments[0].scrollIntoView(false);", down)
    down.click()
    Alert(driver).accept()
    time.sleep(sleep_time)
    Alert(driver).accept()
except:
    print("Error: 送信処理に失敗しました。")

# **📌 スクリプト終了待機**
print("スクリプト完了. Enterキーを押すと終了します...")
input()  # ユーザーが何かキーを押すまで待機
driver.quit()
