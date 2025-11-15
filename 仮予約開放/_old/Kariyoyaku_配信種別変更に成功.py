import win32com.client as win32
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.service import Service
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException

# ── WebDriverのセットアップ ─────────────────────────────
driver_path = EdgeChromiumDriverManager().install()
service = Service(driver_path)
driver = webdriver.Edge(service=service)
driver.maximize_window()

# ── Excelから値を取得 ─────────────────────────────
filename = 'C:/tools/selenium/仮予約開放/仮予約開放.xlsx'
xlApps = win32.Dispatch("Excel.Application")
workbook = xlApps.Workbooks.Open(filename)
sheet = workbook.Worksheets("Sheet1")

def get_excel_value(row, col):
    """Excelのセルの値を取得し、Noneなら空文字を返す"""
    value = sheet.Cells(row, col).Value
    return "" if value is None else str(value).strip()

# 各値を取得（Noneは空文字に変換済み）
edumall_id   = get_excel_value(2, 2)
edumall_pw   = get_excel_value(3, 2)
sellSide_url = get_excel_value(1, 2)
group_id     = get_excel_value(8, 2)
content_id   = get_excel_value(7, 2)
seever_name  = get_excel_value(4, 2)

# ── EduMallにログイン ─────────────────────────────
print("ログイン処理開始...")
driver.get(sellSide_url)
time.sleep(5)
driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[1]/input").send_keys(edumall_id)
driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[2]/input").send_keys(edumall_pw)
driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[3]/button").click()
time.sleep(5)
print("ログイン成功！")

# ── 仮予約画面に遷移 ─────────────────────────────
driver.get("https://school.edumall.jp/dlvr/CAFS15001")
time.sleep(5)

# ── 詳細検索ボタンをクリック ─────────────────────────
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

# ── サーバー名とコンテンツIDを入力 ─────────────────────
print(f"サーバー名: '{seever_name}' を入力中...")
server_name_input = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//input[@data-bind='value: edgeServerName']"))
)
server_name_input.clear()
time.sleep(1)
server_name_input.send_keys(seever_name)
driver.execute_script("""
    arguments[0].value = arguments[1];
    arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
""", server_name_input, seever_name)
time.sleep(1)

print(f"コンテンツID: '{content_id}' を入力中...")
content_id_input = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//input[@data-bind='value: contentsId']"))
)
content_id_input.clear()
time.sleep(1)
content_id_input.send_keys(content_id)
driver.execute_script("""
    arguments[0].value = arguments[1];
    arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
""", content_id_input, content_id)
time.sleep(1)
print("コンテンツIDとサーバー名の入力完了！")

# ── 検索ボタンをクリック ─────────────────────────────
print("検索ボタンを探します...")
try:
    search_button = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//button[contains(text(), '検索')]"))
    )
    driver.execute_script("arguments[0].removeAttribute('disabled');", search_button)
    driver.execute_script("arguments[0].scrollIntoView(true);", search_button)
    time.sleep(1)
    print("検索ボタンをクリックします...")
    driver.execute_script("arguments[0].click();", search_button)
    print("検索実行完了！")
    time.sleep(5)
except Exception as e:
    print(f"検索ボタンの取得またはクリックに失敗しました: {e}")
    driver.quit()
    exit()

# ── 検索結果の1番目の配信IDをクリック ─────────────────
try:
    first_result_span = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//table/tbody/tr[1]/td[2]/span"))
    )
    driver.execute_script("arguments[0].scrollIntoView(true);", first_result_span)
    time.sleep(1)
    driver.execute_script("arguments[0].click();", first_result_span)
    print("検索結果の1番目の配信IDをクリックしました！")
    time.sleep(5)
except Exception as e:
    print(f"検索結果が見つかりませんでした: {e}")
    driver.quit()
    exit()

# ── 詳細画面に遷移したか確認 ─────────────────────────
print("詳細画面に遷移したか確認中...")
if "配信予約詳細" not in driver.title:
    print("詳細画面に遷移できていません。HTMLを確認します。")
    print(driver.page_source)
    driver.quit()
    exit()
print("詳細画面に遷移成功！")

# ── ページ下部までスクロール ─────────────────────────
print("ページを一番下までスクロールします...")
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
time.sleep(2)

# ── 配信予約種別の変更 ─────────────────────────────
print("配信予約種別の変更開始...")

try:
    # <select>要素を取得（SeleniumのSelectクラスを使用）
    haishin_type_box = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//select[contains(@data-bind, 'haishinYoyakuTypeCd')]"))
    )
    
    # disabled属性があれば解除
    if haishin_type_box.get_attribute("disabled"):
        driver.execute_script("arguments[0].removeAttribute('disabled');", haishin_type_box)
        time.sleep(1)
        print("disabled属性を解除しました。")
    
    # SeleniumのSelectクラスで操作
    select = Select(haishin_type_box)
    select.select_by_visible_text("日時")
    time.sleep(1)
    print("配信予約種別を『日時』に変更しました！")
    
    select.select_by_visible_text("通常")
    time.sleep(1)
    print("配信予約種別を『通常』に戻しました！")
    
except Exception as e:
    print(f"配信予約種別の変更に失敗しました: {e}")
    driver.quit()
    exit()

# ── 配信予約ボタンをクリック ─────────────────────────
print("配信予約ボタンを押下します...")
try:
    reserve_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), '配信予約')]"))
    )
    driver.execute_script("arguments[0].scrollIntoView(true);", reserve_button)
    time.sleep(1)
    driver.execute_script("arguments[0].click();", reserve_button)
    print("配信予約ボタンをクリックしました！")
except Exception as e:
    print(f"配信予約ボタンのクリックに失敗しました: {e}")
    driver.quit()
    exit()

# ── ポップアップのOKボタンを待機 ─────────────────────────
print("ポップアップのOKボタンを待機します...")
try:
    popup_ok_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "systemCommonConfirmDialogOk"))
    )
    print("ポップアップのOKボタンが表示されました！")
except Exception as e:
    print(f"ポップアップのOKボタンの待機に失敗しました: {e}")
    driver.quit()
    exit()

# ── 手動確認のためのウエイト ─────────────────────────
input("ポップアップが表示されました。手動で確認後、Enterキーを押してください...")

print("処理完了！ブラウザを閉じます...")
driver.quit()
