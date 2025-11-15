import win32com.client as win32
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
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
filename = r"C:\tools\selenium\SELLSIDE\配信設定ファイル.xlsx"

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
title = str(sheet.Cells(5, 2).Value)  # 登録するタイトル
start_year = str(sheet.Cells(7, 2).Value) # 利用開始日（年）
start_m = str(sheet.Cells(8, 2).Value) # 利用開始日（月）
start_d = str(sheet.Cells(9, 2).Value) # 利用開始日（日）
end_year = str(sheet.Cells(6, 2).Value) # 設定する利用終了日

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
    print("Switched to menu iframe successfully.")
except:
    print("Error: menu iframe not found.")
    exit(1)

# **📌 メニューが閉じている場合は開く**
try:
    menu_toggle = driver.find_element(By.ID, "minimum")
    if menu_toggle.is_displayed():
        menu_toggle.click()
        time.sleep(2)  # メニューが展開されるのを待つ
        print("Clicked #minimum to open menu.")
except:
    print("Error: Could not toggle menu.")

# 6. 「注文管理」をクリック
try:
    order_menu = driver.find_element(By.XPATH, '//p[@onclick="openMenu(\'3\')"]')
    driver.execute_script("arguments[0].click();", order_menu)
    print("注文管理メニュー clicked successfully.")
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
    print("ACCIS注文連携 clicked successfully.")
    time.sleep(1)
except:
    print("Error: ACCIS注文連携 not found.")
    driver.quit()
    exit(1)
    
# **📌 `center` の `iframe` に切り替え**
driver.switch_to.default_content()
try:
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "center")))
    print("Switched to center iframe successfully.")
except:
    print("Error: center iframe not found.")
    exit(1)

# **📌 フォーム入力**
try:
    driver.find_element(By.XPATH, "/html/body/div/form/table[1]/tbody/tr[2]/td[2]/input").send_keys(title)
    driver.find_element(By.XPATH, "/html/body/div/form/table[1]/tbody/tr[4]/td/input[1]").send_keys(start_year)
    driver.find_element(By.XPATH, "/html/body/div/form/table[1]/tbody/tr[4]/td/input[2]").send_keys(start_m)
    driver.find_element(By.XPATH, "/html/body/div/form/table[1]/tbody/tr[4]/td/input[3]").send_keys(start_d)
    driver.find_element(By.XPATH, "/html/body/div/form/table[2]/tbody/tr/td/input[1]").click()
    print("フォーム入力完了.")
except:
    print("Error: フォーム入力に失敗しました.")
    exit(1)

# **📌 無限ループ開始（全て登録すると止まるため問題なし）**
flag = True
loopcounter = 5  # 終了利用期間を変更する用

# `iframe` の切り替え
iframe = driver.find_element(By.XPATH, "/html/body/div/form/iframe")
driver.switch_to.frame(iframe)

while flag:
    try:
        driver.find_element(By.XPATH, "/html/body/form[2]/table[1]/tbody/tr[1]/th[1]/input").click()
        for v in range(20):
            try:
                driver.find_element(By.XPATH, f"/html/body/form[2]/table[1]/tbody/tr[{loopcounter}]/td[4]/input[2]").clear()
                driver.find_element(By.XPATH, f"/html/body/form[2]/table[1]/tbody/tr[{loopcounter}]/td[4]/input[2]").send_keys(end_year)
                loopcounter += 2
            except:
                break

        down = driver.find_element(By.XPATH, "/html/body/form[2]/table[2]/tbody/tr/td[1]/input")
        driver.execute_script("arguments[0].scrollIntoView(false);", down)
        driver.find_element(By.XPATH, "/html/body/form[2]/table[2]/tbody/tr/td[1]/input").click()
        Alert(driver).accept()
        time.sleep(2)
        Alert(driver).accept()
        loopcounter = 5
        time.sleep(1)
    except:
        print("Error: ループ処理中にエラーが発生しました.")
        break

print("スクリプト完了.")
driver.quit()
