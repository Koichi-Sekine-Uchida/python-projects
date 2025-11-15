import sys
import os
import time
import gc
import win32com.client as win32
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# ── Excel解放用の関数 ─────────────────────────
# 今回は設定ファイルを使用しませんが、既存コードに合わせて関数は残します。
def cleanup_excel():
    global xlApps
    try:
        # Excel関連のリソース解放処理はスキップ
        pass 
    except NameError:
        pass
    finally:
        gc.collect()

# ── Edgeのオプション設定 ─────────────────────────
edge_options = EdgeOptions()
edge_options.use_chromium = True
edge_options.add_experimental_option("detach", True)

# ── WebDriverのセットアップ ─────────────────────────
# ★★★ 環境に合わせて、msedgedriver.exeのパスを修正してください ★★★
# 例: EDGE_DRIVER_PATH = "C:\\tools\\python-projects\\EduMall\\msedgedriver\\msedgedriver.exe"
EDGE_DRIVER_PATH = "C:\\tools\\python-projects\\EduMall\\msedgedriver\\msedgedriver.exe" 

try:
    service = Service(EDGE_DRIVER_PATH)
    driver = webdriver.Edge(service=service, options=edge_options)
    driver.maximize_window()
except Exception as e:
    print(f"WebDriverの初期化に失敗しました。Edge WebDriverのパスと設定を確認してください: {e}")
    sys.exit(1)

# ── 設定値の定義（今回はExcel不使用のため直書き） ─────────────────────────
# ログイン情報とURLは環境に合わせて設定してください
EDUMALL_ID   = "sekine@uchida.co.jp" # 適切なIDに修正
EDUMALL_PW   = "e5i3DLU33k"          # 適切なパスワードに修正
TARGET_URL   = "https://school.edumall.jp/home/CABS04005"
SEARCH_SCHOOL_NAME = "【不使用】" # 今回固定の検索キーワード

# L-Gate連携ソースIDが一覧の8列目にあると推定
L_GATE_ID_COLUMN_INDEX = 8 

# ── EduMallにログイン ─────────────────────────
print("ログイン処理開始...")
try:
    driver.get(TARGET_URL)
    
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, "j_username"))
    )
    driver.find_element(By.NAME, "j_username").send_keys(EDUMALL_ID)
    driver.find_element(By.NAME, "j_password").send_keys(EDUMALL_PW)
    driver.find_element(By.ID, "login_button").click()
    time.sleep(5)
    print("ログイン成功！")
except Exception as e:
    print(f"ログイン処理中にエラー: {e}")
    cleanup_excel()
    sys.exit()

# ── 検索画面への遷移と検索実行 ─────────────────────────

try:
    # 検索画面の学校名入力フィールドを待つ（form:schoolName）
    school_name_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, "form:schoolName"))
    )
    school_name_input.clear()
    school_name_input.send_keys(SEARCH_SCHOOL_NAME)
    
    # 検索ボタンをクリック（form:search）
    search_button = driver.find_element(By.ID, "form:search")
    search_button.click()
    time.sleep(3) 
    print(f"学校名「{SEARCH_SCHOOL_NAME}」で検索を実行しました。")

    # 検索結果の一覧表示を待つ
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "form:schoolListTable"))
    )
    
    # ── L-Gate連携ソースIDのヘッダーをクリックしてソート ──────────────────
    # データがあるものだけを上部に集める
    l_gate_header_xpath = f"//table[@id='form:schoolListTable']/thead/tr/th[{L_GATE_ID_COLUMN_INDEX}]/a"
    l_gate_header = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, l_gate_header_xpath))
    )
    # 昇順ソート（空欄が後に来る）になるようクリック
    l_gate_header.click() 
    print("L-Gate連携ソースIDのヘッダーをクリックしてソートしました。")
    time.sleep(3) 

except Exception as e:
    print(f"検索またはソート処理に失敗: {e}")
    sys.exit(1)


# ── 総ページ数を取得 ─────────────────────────
total_pages = 1
try:
    page_links = driver.find_elements(By.XPATH, "//a[@class='page-link']/span[text()!='...']")
    page_numbers = [int(link.text) for link in page_links if link.text.isdigit()]
    if page_numbers:
        total_pages = max(page_numbers)
except Exception:
    total_pages = 1

print(f"総ページ数: {total_pages}")


# ── ページを順番に処理する関数 ─────────────────────────
processed_count = 0
def process_rows_on_current_page():
    global processed_count
    
    # L-Gate IDが存在する行があったかどうか
    is_data_present = False
    
    try:
        # 一覧のtbody内の全行を取得
        rows = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located(
                (By.XPATH, "//table[@id='form:schoolListTable']/tbody/tr")
            )
        )
    except Exception as e:
        print(f"学校一覧の取得に失敗しました: {e}")
        return is_data_present 

    row_count = len(rows)
    print(f"  - このページの行数: {row_count}")

    # 行のインデックスを使って処理（DOM変化対策）
    for i in range(1, row_count + 1):
        
        current_row_xpath = f"//table[@id='form:schoolListTable']/tbody/tr[{i}]"
        
        try:
            # L-Gate連携ソースIDの列（td要素）を取得 
            l_gate_id_cell_xpath = f"{current_row_xpath}/td[{L_GATE_ID_COLUMN_INDEX}]"
            l_gate_id_cell = driver.find_element(By.XPATH, l_gate_id_cell_xpath)
            l_gate_id_value = l_gate_id_cell.text.strip()

            school_name_xpath = f"{current_row_xpath}/td[2]" 
            school_name = driver.find_element(By.XPATH, school_name_xpath).text.strip()
            
        except NoSuchElementException:
            print(f"  - 行 {i} の要素が見つかりませんでした。スキップします。")
            continue
        except Exception as e:
            print(f"  - 行 {i} の情報取得中にエラー: {e}")
            continue

        
        # L-Gate連携ソースIDが設定されているかチェック
        if l_gate_id_value:
            is_data_present = True # データが存在する行があった
            print(f"  - 学校名: {school_name} (ID: {l_gate_id_value}) -> 処理対象です。")

            # ── 詳細ボタンをクリックして詳細画面へ遷移 ─────────────────────────
            try:
                # 1列目にある詳細ボタン（aタグ）をクリック
                # ご指摘の通り、SCから始まる学校IDの列ではなく、詳細ボタンの列です
                detail_button_xpath = f"{current_row_xpath}/td[1]/a" 
                detail_button = driver.find_element(By.XPATH, detail_button_xpath)
                driver.execute_script("arguments[0].scrollIntoView(true);", detail_button)
                time.sleep(1)
                detail_button.click()
                
                # 詳細画面の表示を待つ (L-Gate ID入力フィールドの名前を推定: form:lGateSourceId)
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.NAME, "form:lGateSourceId"))
                )
                print("  - 詳細画面に遷移しました。")
            except Exception as e:
                print(f"  - 詳細画面への遷移に失敗: {e}。一覧に戻ります。")
                driver.get(TARGET_URL) 
                continue 

            
            # ── L-Gate連携ソースIDを空欄にして更新 ─────────────────────────
            try:
                # L-Gate連携ソースIDの入力フィールドを特定
                l_gate_input = driver.find_element(By.NAME, "form:lGateSourceId")
                
                # 値をクリア
                l_gate_input.clear()
                print("  - L-Gate連携ソースIDを空欄に設定しました。")
                
                # 更新ボタンをクリック（IDを推定: form:update）
                update_button = driver.find_element(By.ID, "form:update")
                driver.execute_script("arguments[0].scrollIntoView(true);", update_button)
                time.sleep(1)
                update_button.click()
                
                # 更新完了後のメッセージまたは一覧画面への遷移を待つ
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "form:schoolListTable"))
                )
                
                processed_count += 1
                print(f"  - {school_name} のL-Gate連携ソースIDを空欄にし、更新しました。")

            except Exception as e:
                print(f"  - 更新処理中にエラーが発生: {e}。一覧に戻ります。")
                # エラー時の復帰処理
                try:
                    back_button = driver.find_element(By.ID, "form:back")
                    back_button.click()
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "form:schoolListTable"))
                    )
                except:
                    driver.get(TARGET_URL)
                    
                continue 

        else:
            # L-Gate連携ソースIDが空欄の場合
            print(f"  - 学校名: {school_name} -> L-Gate連携ソースIDが空欄です。")
            # ★ソート済みなので、空欄の行が出てきたら処理を終了します★
            return is_data_present

    return is_data_present


# ── 全ページを順番に処理 ─────────────────────────
for p in range(1, total_pages + 1):
    print(f"\n=== ページ {p} を処理します ===")

    # ページ移動のロジック
    if p > 1:
        try:
            page_link_xpath = f"//a[@class='page-link']/span[text()='{p}']"
            page_link_elem = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, page_link_xpath))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", page_link_elem)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", page_link_elem)
            time.sleep(5) 
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "form:schoolListTable"))
            )
        except Exception as e:
            print(f"{p}ページ目への移動または一覧表示待機に失敗しました: {e}")
            break

    # ── 現在のページの行を処理 ─────────────────
    data_was_present = process_rows_on_current_page()
    
    # L-Gate連携ソースIDが設定されている行がなければ、処理を終了
    if not data_was_present:
        print("\nソートされた一覧の途中でL-Gate連携ソースIDが空欄の行が見つかりました。以降のレコードも空欄と判断し、処理を終了します。")
        break


# ── 処理完了 ─────────────────────────
print(f"\n==========================================")
print(f"すべての処理が完了しました。")
print(f"合計 {processed_count} 件のL-Gate連携ソースIDを空欄にしました。")
print(f"==========================================")

cleanup_excel()