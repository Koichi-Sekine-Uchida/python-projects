import win32com.client as win32
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.service import Service
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException

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
def get_excel_value(row, col):
    """Excelのセルの値を取得し、Noneなら空文字を返す"""
    value = sheet.Cells(row, col).Value
    return "" if value is None else str(value).strip()

# **取得データ**
edumall_id = get_excel_value(2, 2)
edumall_pw = get_excel_value(3, 2)
sellSide_url = get_excel_value(1, 2)
group_id = get_excel_value(8, 2)
content_id = get_excel_value(7, 2)  # **Noneを空文字に変換済み**
seever_name = get_excel_value(4, 2)  # **Noneを空文字に変換済み**

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
try:
    detail_search_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//label[contains(text(), '詳細検索')]"))
    )
    driver.execute_script("arguments[0].click();", detail_search_button)
    time.sleep(2)
except TimeoutException:
    print("詳細検索ボタンが見つかりませんでした")
    driver.quit()
    exit()

# **サーバー名を入力**
print(f"サーバー名: '{seever_name}' を入力中...")
server_name_input = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//input[@data-bind='value: edgeServerName']"))
)
server_name_input.clear()
time.sleep(1)
server_name_input.send_keys(seever_name)
time.sleep(1)

# **JavaScriptを使って値を確実に入力**
driver.execute_script("""
    arguments[0].value = arguments[1];
    arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
""", server_name_input, seever_name)
time.sleep(1)

# **コンテンツIDを入力**
print(f"コンテンツID: '{content_id}' を入力中...")
content_id_input = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//input[@data-bind='value: contentsId']"))
)
content_id_input.clear()
time.sleep(1)
content_id_input.send_keys(content_id)
time.sleep(1)

# **JavaScriptを使って値を確実に入力**
driver.execute_script("""
    arguments[0].value = arguments[1];
    arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
""", content_id_input, content_id)
time.sleep(1)

print("コンテンツIDの入力完了！")
print("サーバー名の入力完了！")

# **検索ボタンをクリック**
print("検索ボタンを探します...")
try:
    search_button = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//button[contains(text(), '検索')]"))
    )
    driver.execute_script("arguments[0].click();", search_button)
    print("検索実行完了！")
    time.sleep(5)
except Exception as e:
    print(f"検索ボタンの取得またはクリックに失敗しました: {e}")
    driver.quit()
    exit()

# **検索結果の1番目をクリック**
try:
    first_result_span = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//table/tbody/tr[1]/td[2]/span"))
    )
    driver.execute_script("arguments[0].click();", first_result_span)
    print("検索結果の1番目の配信IDをクリックしました！")
    time.sleep(5)
except Exception as e:
    print(f"検索結果が見つかりませんでした: {e}")
    driver.quit()
    exit()

# **ページを一番下までスクロール**
print("ページを一番下までスクロールします...")
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
time.sleep(2)  # スクロール完了待ち

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

# **配信予約種別の変更処理**
print("配信予約種別の変更開始...")

try:
    # **リストボックスの要素を取得**
    haishin_type_box = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//select[@data-bind='value: haishinYoyakuTypeCd']"))
    )

    # **JavaScriptでドロップダウンを開く**
    driver.execute_script("arguments[0].focus();", haishin_type_box)
    driver.execute_script("arguments[0].click();", haishin_type_box)
    time.sleep(1)

    print("リストボックスの枠をクリックしました...")

    # **「日時」のオプションをクリック**
    date_option = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//option[contains(text(), '日時')]"))
    )

    # **オプションを選択**
    date_option.click()
    time.sleep(1)

    print("配信予約種別を「日時」に変更しました！")

except Exception as e:
    print(f"配信予約種別の変更に失敗しました: {e}")


print("処理完了！ブラウザを閉じます...")
driver.quit()
