import win32com.client as win32
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
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
edumall_id = str(sheet.Cells(2, 2).Value)    # EduMallのID
edumall_pw = str(sheet.Cells(3, 2).Value)    # EduMallのPW

# 1. アクセスするリンク
driver.get(sellSide_url)
time.sleep(3)

# 2. ログイン処理
driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[1]/input").send_keys(edumall_id + Keys.TAB)
driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[2]/input").send_keys(edumall_pw)
driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[3]/button").click()
time.sleep(3)

# 3. `menu` iframe に切り替え
try:
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "menu")))
    driver.switch_to.frame("menu")
    print("Switched to menu iframe successfully.")
except:
    print("Error: menu iframe not found.")
    driver.quit()
    exit(1)

# 4. #minimum をクリックしてメニューを展開
try:
    # "minimum" 内の <td> をクリック (A: 直接IDをクリックしてもOK)
    td_in_minimum = driver.find_element(By.XPATH, '//*[@id="minimum"]//td')
    driver.execute_script("arguments[0].click();", td_in_minimum)
    print("Clicked #minimum to open menu.")
except:
    print("Error: #minimum not found or not clickable.")
    driver.quit()
    exit(1)

# 5. #menu-wrap が表示されるのを待つ
try:
    WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, "menu-wrap"))
    )
    print("#menu-wrap is now visible.")
except:
    print("Error: #menu-wrap did not become visible.")
    driver.quit()
    exit(1)

time.sleep(2)  # 念のため

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
except:
    print("Error: ACCIS注文連携 not found.")
    driver.quit()
    exit(1)

print("Script executed successfully.")
# driver.quit()
