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

# ── WebDriverのセットアップ ─────────────────────────────
driver_path = EdgeChromiumDriverManager().install()
service = Service(driver_path)
driver = webdriver.Edge(service=service)
driver.maximize_window()

# ── Excelファイルのパス（Pythonファイルと同じフォルダの場合） ──
current_dir = os.path.dirname(os.path.abspath(__file__))
excel_path = os.path.join(current_dir, "仮予約開放.xlsx")

# ── Excelから値を取得 ─────────────────────────────
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
group_id     = get_excel_value(8, 2)
content_id   = get_excel_value(7, 2)
seever_name  = get_excel_value(4, 2)
start_date   = get_excel_value(5, 2)  # 配信予定日時（開始）
end_date     = get_excel_value(6, 2)  # 配信予定日時（終了）

# ── 処理を繰り返す回数をユーザーに入力 ─────────────────────────────
try:
    cycles = int(input("処理を繰り返す回数を入力してください: "))
except Exception as e:
    print(f"回数の入力エラー: {e}")
    driver.quit()
    exit()

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

# ── 更新予定日時の入力（配信予定日時欄） ─────────────────────
print(f"配信予定日時(開始): '{start_date}' を入力中...")
start_date_input = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//input[@data-bind='value: haishinYoteiDatetimeFrom']"))
)
start_date_input.clear()
time.sleep(1)
start_date_input.send_keys(start_date)
driver.execute_script("""
    arguments[0].value = arguments[1];
    arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
""", start_date_input, start_date)
time.sleep(1)
print("配信予定日時(開始)の入力完了！")

print(f"配信予定日時(終了): '{end_date}' を入力中...")
end_date_input = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//input[@data-bind='value: haishinYoteiDatetimeTo']"))
)
end_date_input.clear()
time.sleep(1)
end_date_input.send_keys(end_date)
driver.execute_script("""
    arguments[0].value = arguments[1];
    arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
""", end_date_input, end_date)
time.sleep(1)
print("配信予定日時(終了)の入力完了！")

# ── 表示件数を100に変更 ─────────────────────────────
try:
    display_count_select = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//select[contains(@data-bind, 'displayCount')]"))
    )
    Select(display_count_select).select_by_value("100")
    driver.execute_script("arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", display_count_select)
    time.sleep(1)
    print("表示件数を 100 に変更しました！")
except Exception as e:
    print(f"表示件数の変更に失敗しました: {e}")

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

# ── 繰り返し処理 ─────────────────────────
for i in range(cycles):
    print(f"【繰り返し {i+1} 回目】")
    # ── 配信予約種別の変更 ─────────────────────────
    try:
        haishin_type_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//select[contains(@data-bind, 'haishinYoyakuTypeCd')]"))
        )
        if haishin_type_box.get_attribute("disabled"):
            driver.execute_script("arguments[0].removeAttribute('disabled');", haishin_type_box)
            time.sleep(1)
            print("disabled属性を解除しました。")
        # SeleniumのSelectで変更（内部値「01」=「日時」、 「03」=「通常」と仮定）
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
        
    # ── ポップアップの内容をチェックし、容量不足の場合は送信をスキップ ─────────────────────────
    print("ポップアップの内容を確認します...")
    check_Text = "ディスク容量"  # この単語が含まれている場合、送信はスキップします。
    try:
        text_element = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[4]/div/div/div[2]/p"))
        )
        text_content = text_element.text
        print(f"ポップアップのメッセージ: {text_content}")
        if check_Text in text_content:
            print("警告: 容量不足のため、送信をスキップします。")
            cancel_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div[4]/div/div/div[3]/div/div[1]/button"))
            )
            driver.execute_script("arguments[0].click();", cancel_button)
            time.sleep(2)
            # 一覧画面に戻る処理
            back_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/section/div/div/div[8]/div[1]/a"))
            )
            driver.execute_script("arguments[0].click();", back_button)
            time.sleep(3)
        else:
            print("容量不足ではないため、送信を継続します。")
            confirm_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div[4]/div/div/div[3]/div/div[2]/button"))
            )
            driver.execute_script("arguments[0].click();", confirm_button)
            time.sleep(3)
    except TimeoutException:
        print("ポップアップが表示されませんでした。処理を継続します。")
    except Exception as e:
        print(f"ポップアップ処理に失敗しました: {e}")
        driver.quit()
        exit()
        
    # ── 一覧画面に戻った後、リストの先頭の配信IDをクリック ─────────────────────────
    print("一覧に戻ったので、先頭の配信IDをクリックします...")
    try:
        # ※一覧画面に戻ったことをテーブルの存在で判断
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//table/tbody/tr"))
        )
        first_result_span = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//table/tbody/tr[1]/td[2]/span"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", first_result_span)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", first_result_span)
        print("一覧画面の先頭の配信IDをクリックしました！")
        time.sleep(5)
    except Exception as e:
        print(f"一覧へ戻る（先頭リンククリック）の処理に失敗しました: {e}")
        driver.quit()
        exit()

# ── Excelファイルを閉じる（変更を保存しない場合は False） ─────────────────────────
workbook.Close(False)
xlApps.Quit()
del sheet, workbook, xlApps

print("全処理完了！ブラウザを閉じます...")
driver.quit()
