import sys
import os
import time
import gc
import win32com.client as win32
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options as EdgeOptions

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
gakkou_name  = get_excel_value(4, 2)
kokyaku_name = get_excel_value(5, 2)

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


# ── 「学校名」を入力 (修正版) ─────────────────────────
if sheet.Cells(4, 2).value is not None:
    gakkou_name_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//input[@data-bind='value:gakkoName']"))
    )
    gakkou_name_input.clear()
    driver.execute_script("""
        arguments[0].value = arguments[1];
        arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
        arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
    """, gakkou_name_input, gakkou_name)

# ── 「顧客名」を入力 (修正版) ─────────────────────────
if sheet.Cells(5, 2).value is not None:
    kokyaku_name_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//input[@data-bind='value:kokyakuName']"))
    )
    kokyaku_name_input.clear()
    driver.execute_script("""
        arguments[0].value = arguments[1];
        arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
        arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
    """, kokyaku_name_input, kokyaku_name)

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
    time.sleep(10)
except Exception as e:
    print(f"検索ボタンの取得またはクリックに失敗しました: {e}")
    cleanup_excel()
    sys.exit()

# ここから先、複数ページがある場合に対応するための処理

def process_rows_on_current_page():
    """現在表示中のページ内の学校行を順番に処理する関数"""
    try:
        rows = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located(
                (By.XPATH, "//*[@id='mainContens']/div[3]/div/datatable/div[2]/div/table/tbody/tr")
            )
        )
    except Exception as e:
        print(f"学校一覧の取得に失敗しました: {e}")
        return  # 何もせず戻る

    row_count = len(rows)
    print(f"このページの行数: {row_count}")

    for i in range(1, row_count + 1):
        print(f"【学校 {i} の処理開始】")
        try:
            # テーブル行を再取得（DOM変化対策）
            rows = driver.find_elements(
                By.XPATH,
                "//*[@id='mainContens']/div[3]/div/datatable/div[2]/div/table/tbody/tr"
            )
            school_id_xpath = f"//*[@id='mainContens']/div[3]/div/datatable/div[2]/div/table/tbody/tr[{i}]/td[1]/span"
            school_id_elem = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, school_id_xpath))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", school_id_elem)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", school_id_elem)
            print(f"行 {i} の学校IDをクリックしました。")
            time.sleep(3)  # 詳細画面への遷移待ち
        except Exception as e:
            print(f"行 {i} の学校IDのクリックに失敗しました: {e}")
            continue

        # 学校名フィールドを【不使用】に更新
        try:
            gakko_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//*[@id='mainContens']/div[5]/div[2]/div[4]/div[1]/input")
                )
            )
            current_value = gakko_input.get_attribute("value")
            if not current_value.startswith("【不使用】"):
                new_value = "【不使用】" + current_value
            else:
                new_value = current_value

            driver.execute_script("""
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            """, gakko_input, new_value)
            print(f"学校名を更新しました: {new_value}")
        except Exception as e:
            print(f"学校名入力フィールドの更新に失敗しました: {e}")
            # 詳細画面から一覧に戻る
            time.sleep(2)
            driver.get("https://school.edumall.jp/schl/CAAS11001")
            time.sleep(5)
            continue

        # 更新ボタンをクリック
        try:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1)
            update_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//*[@id='mainContens']/div[12]/div[2]/button")
                )
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", update_button)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", update_button)
            print("更新ボタンをクリックしました。")
            time.sleep(2)
        except Exception as e:
            print(f"更新ボタンのクリックに失敗しました: {e}")
            time.sleep(2)
            driver.get("https://school.edumall.jp/schl/CAAS11001")
            time.sleep(5)
            continue

        # ポップアップのOKボタンをクリック
        try:
            popup_ok_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'OK')]"))
            )
            popup_ok_button.click()
            print("ポップアップのOKボタンをクリックしました。")
            time.sleep(3)
        except:
            # ポップアップが出ない場合は何もしない
            pass

        # 一覧画面に自動的に戻っている想定なら、ロード完了を待機
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//*[@id='mainContens']/div[3]/div/datatable/div[2]/div/table/tbody/tr")
                )
            )
            print("一覧画面に戻りました。")
            time.sleep(3)
        except Exception as e:
            print(f"一覧画面の確認に失敗しました: {e}")
            # もし戻っていない場合は直接URLへ飛ぶ
            driver.get("https://school.edumall.jp/schl/CAAS11001")
            time.sleep(5)

        print(f"【学校 {i} の処理完了】\n")


# ── 総ページ数を取得 ─────────────────────────
# 例：下記のようなページリンクがある場合
# <a href="#" class="page-link" data-bind="..."><span data-bind="text: pageIndex">1</span></a>
# <a href="#" class="page-link" data-bind="..."><span data-bind="text: pageIndex">2</span></a>
# ...
# 「前へ」「次へ」などのリンクを含む場合は除外処理が必要
try:
    page_spans = driver.find_elements(By.XPATH, "//a[@class='page-link']/span")
    # ページ番号だけを取得（数字のみ）
    page_numbers = []
    for span in page_spans:
        txt = span.text.strip()
        if txt.isdigit():
            page_numbers.append(int(txt))
    if not page_numbers:
        # ページリンクが無い or 1ページのみの場合
        total_pages = 1
    else:
        total_pages = max(page_numbers)
except Exception as e:
    print(f"ページ数の取得に失敗: {e}")
    total_pages = 1

print(f"総ページ数: {total_pages}")

# ── ページを順番に処理 ─────────────────────────
for p in range(1, total_pages + 1):
    print(f"\n=== ページ {p} を処理します ===")

    # ページ1以外の場合、該当ページリンクをクリックして移動
    # ただし、既にページpに居る場合はスキップ
    if p > 1:
        # ページリンクのspan要素を検索
        try:
            # 「p」ページへのリンクを探す
            page_link_xpath = f"//a[@class='page-link']/span[text()='{p}']"
            page_link_elem = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, page_link_xpath))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", page_link_elem)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", page_link_elem)
            time.sleep(3)
        except Exception as e:
            print(f"{p}ページ目への移動に失敗しました: {e}")
            break

        # ページ遷移後の一覧読み込みを待機
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//*[@id='mainContens']/div[3]/div/datatable/div[2]/div/table/tbody/tr")
                )
            )
            time.sleep(3)
        except Exception as e:
            print(f"{p}ページ目の一覧読み込みに失敗しました: {e}")
            break

    # ── 現在のページの行を処理 ─────────────────
    process_rows_on_current_page()

print("すべてのページの更新処理が完了しました。")
cleanup_excel()
