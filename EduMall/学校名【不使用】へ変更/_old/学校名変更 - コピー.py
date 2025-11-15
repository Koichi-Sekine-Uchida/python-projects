import sys
import win32com.client as win32
import time, os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.service import Service
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.edge.options import Options as EdgeOptions
import gc
# ── Excel解放用の関数 ─────────────────────────
def cleanup_excel():
    global sheet, workbook, xlApps
    try:
        workbook.Close(False)
        xlApps.Quit()
    except Exception as e:
        print(f"Excelのクローズ処理中にエラー: {e}")
    finally:
        try:
            del sheet, workbook, xlApps
        except Exception as e:
            print(f"変数削除中にエラー: {e}")
        gc.collect()
        print("Excelのリソースを解放しました。")

# ── Edgeのオプション設定（detachを有効化） ─────────────────────────
edge_options = EdgeOptions()
edge_options.use_chromium = True
edge_options.add_experimental_option("detach", True)

# ── WebDriverのセットアップ ─────────────────────────
driver_path = EdgeChromiumDriverManager().install()
service = Service(driver_path)
driver = webdriver.Edge(service=service, options=edge_options)
driver.maximize_window()

# ── Excelファイルのパス（Pythonファイルと同じフォルダの場合） ─────────────────
current_dir = os.path.dirname(os.path.abspath(__file__))
excel_path = os.path.join(current_dir, "設定.xlsx")

# ── Excelから値を取得 ─────────────────────────
xlApps = win32.Dispatch("Excel.Application")
workbook = xlApps.Workbooks.Open(excel_path)
sheet = workbook.Worksheets("Sheet1")

def get_excel_value(row, col):
    """Excelのセルの値を取得し、Noneなら空文字を返す"""
    value = sheet.Cells(row, col).Value
    return "" if value is None else str(value).strip()

# 各値を取得（Noneは空文字に変換済み）
edumall_id   = get_excel_value(2, 2)
edumall_pw   = get_excel_value(3, 2)
sellSide_url = get_excel_value(1, 2)
gakkou_name = str(sheet.Cells(4,2))
kokyaku_name = str(sheet.Cells(5,2))


# ── EduMallにログイン ─────────────────────────
print("ログイン処理開始...")
driver.get(sellSide_url)
time.sleep(5)
try:
    driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[1]/input").send_keys(edumall_id)
    driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[2]/input").send_keys(edumall_pw)
    driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[3]/button").click()
except Exception as e:
    print(f"ログイン処理中にエラー: {e}")
    cleanup_excel()
    sys.exit()
time.sleep(5)
print("ログイン成功！")

# ── 学校画面に遷移 ─────────────────────────
driver.get("https://school.edumall.jp/schl/CAAS11001")
time.sleep(5)


# 「学校名」を入力
if sheet.Cells(4,2).value is not None:
    gakkou_name_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//input[@data-bind='value:gakkoName']"))
    )
    gakkou_name_input.clear()
    gakkou_name_input.send_keys(gakkou_name)

# 「顧客名」を入力
if sheet.Cells(5,2).value is not None:
    kokyaku_name_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//input[@data-bind='value:kokyakuName']"))
    )
    kokyaku_name_input.clear()
    kokyaku_name_input.send_keys(kokyaku_name)


time.sleep(2)  # 展開が完了するのを待つ

# ── 検索ボタンをクリック ─────────────────────────
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
    cleanup_excel()
    sys.exit()

# ── 119行目以降：学校一覧の各行に対する処理 ─────────────────────────

# まず、一覧に表示されている学校の行数を取得
try:
    rows = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.XPATH, "//*[@id='mainContens']/div[3]/div/datatable/div[2]/div/table/tbody/tr"))
    )
    row_count = len(rows)
    print(f"一覧にある学校の行数: {row_count}")
except Exception as e:
    print(f"学校一覧の取得に失敗しました: {e}")
    cleanup_excel()
    sys.exit()

# 各学校行に対して処理を実施（1行目から順に）
for i in range(1, row_count + 1):
    try:
        # 学校IDをクリック（XPathは各行毎にインデックスを付与）
        row_xpath = f"//*[@id='mainContens']/div[3]/div/datatable/div[2]/div/table/tbody/tr[{i}]/td[1]/span"
        school_id_elem = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, row_xpath))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", school_id_elem)
        time.sleep(1)
        school_id_elem.click()
        print(f"行 {i} の学校IDをクリックしました。")
        time.sleep(5)  # 次画面遷移待ち
    except Exception as e:
        print(f"行 {i} の学校IDのクリックに失敗しました: {e}")
        continue  # 次の行へ進む

# ----- 学校IDをクリック後の処理例 -----

# 学校名入力フィールドの更新（既存の値の先頭に【不使用】を追加）
try:
    gakko_name_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='mainContens']/div[5]/div[2]/div[4]/div[1]/input"))
    )
    current_value = gakko_name_input.get_attribute("value")
    new_value = "【不使用】" + current_value
    # JavaScriptで値を直接更新（clearせず上書き）
    driver.execute_script("arguments[0].value = arguments[1];", gakko_name_input, new_value)
    print(f"学校名に【不使用】を追加しました: {new_value}")
except Exception as e:
    print(f"学校名入力フィールドの更新に失敗しました: {e}")
    driver.back()
    time.sleep(5)
    # 次の行へ進む場合はcontinueなど適切に処理してください

# 画面下部までスクロールして更新ボタンをクリック
try:
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(1)
    update_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//*[@id='mainContens']/div[12]/div[2]/button"))
    )
    driver.execute_script("arguments[0].scrollIntoView(true);", update_button)
    time.sleep(1)
    update_button.click()
    print("更新ボタンをクリックしました。")
    time.sleep(2)  # 更新処理待ち
except Exception as e:
    print(f"更新ボタンのクリックに失敗しました: {e}")
    driver.back()
    time.sleep(5)
    # 次の行へ進む場合はcontinueなど適切に処理してください

# 更新ボタン押下後に表示されるポップアップのOKボタンをクリック
try:
    popup_ok_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'OK')]"))
    )
    popup_ok_button.click()
    print("ポップアップのOKボタンをクリックしました。")
    time.sleep(2)
except Exception as e:
    print(f"ポップアップのOKボタンのクリックに失敗しました: {e}")

# 更新後、詳細画面の学校名入力フィールドから値を取得する
try:
    updated_gakko_name_elem = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='mainContens']/div[5]/div[2]/div[4]/div[1]/input"))
    )
    updated_gakko_name = updated_gakko_name_elem.get_attribute("value")
    print("更新後の学校名:", updated_gakko_name)
except Exception as e:
    print(f"詳細画面の学校名の取得に失敗しました: {e}")
