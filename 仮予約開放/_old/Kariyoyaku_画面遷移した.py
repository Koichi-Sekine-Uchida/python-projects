import win32com.client as win32
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.service import Service

# **WebDriverのセットアップ**
driver_path = EdgeChromiumDriverManager().install()
service = Service(driver_path)
driver = webdriver.Edge(service=service)

# **ブラウザを最大化**
driver.maximize_window()

# **Excelを開く**
filename = 'C:/tools/selenium/仮予約開放/仮予約開放.xlsx'
xlApps = win32.Dispatch("Excel.Application")
workbook = xlApps.Workbooks.Open(filename)
sheet = workbook.Worksheets("Sheet1")

# **Excelからデータ取得**
edumall_id = str(sheet.Cells(2, 2).Value)
edumall_pw = str(sheet.Cells(3, 2).Value)
sellSide_url = str(sheet.Cells(1, 2).Value)
group_id = str(sheet.Cells(8, 2).Value)
content_id = str(sheet.Cells(7, 2).Value)
server_name = str(sheet.Cells(4, 2).Value)  # エッジサーバ名の取得

# **EduMallにログイン**
print("ログイン処理開始...")
driver.get(sellSide_url)
time.sleep(5)
driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[1]/input").send_keys(edumall_id)
driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[2]/input").send_keys(edumall_pw)
driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[3]/button").click()
time.sleep(5)
print("ログイン成功！")

# **仮予約画面に遷移**
driver.get("https://school.edumall.jp/dlvr/CAFS15001")
time.sleep(5)

# **詳細検索ボタンをクリック**
print("詳細検索ボタンをクリックします...")
detail_search_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//label[contains(text(), '詳細検索')]"))
)
detail_search_button.click()
time.sleep(2)  # 展開を待つ

# **コンテンツIDを入力**
print(f"コンテンツID {content_id} を入力中...")
content_id_input = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//input[@data-bind='value: contentsId']"))
)
content_id_input.clear()
content_id_input.send_keys(content_id)
print("コンテンツIDの入力完了！")

# **エッジサーバ名を入力**
print(f"エッジサーバ名 {server_name} を入力中...")
server_name_input = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//input[@data-bind='value: edgeServerName']"))
)
server_name_input.clear()
server_name_input.send_keys(server_name)
print("エッジサーバ名の入力完了！")

# **検索ボタンをクリック**
print("検索ボタンをクリックします...")
search_button = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//button[contains(text(), '検索')]"))
)

# JavaScriptで検索ボタンを有効化（もし無効なら）
driver.execute_script("arguments[0].removeAttribute('disabled');", search_button)
driver.execute_script("arguments[0].scrollIntoView(true);", search_button)
time.sleep(1)

# クリックを試みる
try:
    search_button.click()
    print("通常クリックで検索ボタンを押しました。")
except:
    driver.execute_script("arguments[0].click();", search_button)
    print("JavaScriptで検索ボタンを押しました。")

print("検索実行完了！ 検索結果が表示されるのを待っています...")

# **検索結果が表示されるのを待つ**
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//table/tbody/tr"))
)
print("検索結果が表示されました！")

# **検索結果の配信IDをクリック**
try:
    # まずテーブルが存在するか確認
    results_table = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//table/tbody"))
    )
    print("検索結果のテーブルが検出されました。")

    # **配信IDの `span` タグを取得**
    first_result_span = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//table/tbody/tr[1]/td[2]/span"))
    )
    print("検索結果の配信IDの `span` タグを取得しました。")

    # **クリックが必要な場合は JavaScript を使用**
    driver.execute_script("arguments[0].scrollIntoView(true);", first_result_span)
    time.sleep(1)
    driver.execute_script("arguments[0].click();", first_result_span)
    print("検索結果の1番目の配信IDをクリックしました！")
    time.sleep(5)
except Exception as e:
    print(f"エラー発生: {e}")
    print("検索結果が見つかりませんでした。処理を終了します。")
    driver.quit()
    exit()
