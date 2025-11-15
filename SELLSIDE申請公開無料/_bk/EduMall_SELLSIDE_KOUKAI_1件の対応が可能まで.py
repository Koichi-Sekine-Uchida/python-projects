import pythoncom
import win32com.client
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

# ==============================================
# ※ 事前に以下のコマンドを実行してください
# 管理者権限のコマンドプロンプトまたは PowerShell で
#   python -m win32com.client.makepy
# を実行し、Excel用のライブラリ（例: Microsoft Excel 16.0 Object Library）を選択してください。
# ==============================================

# ==============================================
# Edgeドライバーのインストール
# ==============================================
driver_path = EdgeChromiumDriverManager().install()
service = Service(driver_path)
driver = webdriver.Edge(service=service)

# ==============================================
# Excelファイルから必要情報を取得
# ==============================================
filename = r"C:\tools\python-projects\SELLSIDE申請公開\配信設定ファイル.xlsx"

if not os.path.exists(filename):
    print(f"エラー: 指定されたExcelファイルが見つかりません: {filename}")
    driver.quit()
    exit(1)

try:
    pythoncom.CoInitialize()
    excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    workbook = excel.Workbooks.Open(filename)
    sheet = workbook.Worksheets("Sheet1")
    sellSide_url = str(sheet.Cells(1, 2).Value)
    edumall_id   = str(sheet.Cells(2, 2).Value)
    edumall_pw   = str(sheet.Cells(3, 2).Value)
    sleep_time   = int(sheet.Cells(4, 2).Value)
    content_id   = str(sheet.Cells(2, 5).Value)
    workbook.Close(SaveChanges=False)
    excel.Quit()
except Exception as e:
    print("Excelファイルの読み込み時にエラーが発生しました:", e)
    driver.quit()
    exit(1)

# ==============================================
# Webアプリへアクセス & ログイン
# ==============================================
try:
    driver.get(sellSide_url)
    time.sleep(3)
    driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[1]/input").send_keys(edumall_id + Keys.TAB)
    driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[2]/input").send_keys(edumall_pw)
    driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[3]/button").click()
except Exception as e:
    print("ログイン処理中にエラーが発生しました:", e)
    driver.quit()
    exit(1)

# ==============================================
# メニュー操作（menu iframe へ切り替え）
# ==============================================
try:
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "menu")))
    driver.switch_to.frame("menu")
except Exception as e:
    print("Error: menu iframe が見つかりません。", e)
    driver.quit()
    exit(1)

try:
    contents_menu = driver.find_element(By.XPATH, '//p[@onclick="openMenu(\'1\')"]')
    driver.execute_script("arguments[0].click();", contents_menu)
    time.sleep(1)
except Exception as e:
    print("Error: コンテンツ管理メニューが見つかりません。", e)
    driver.quit()
    exit(1)

try:
    contents_regist_link = driver.find_element(
        By.XPATH,
        '//a[@onclick="showPage(this, \'goods/CGdGoodsSearch.jsp\'); return false;"]'
    )
    driver.execute_script("arguments[0].click();", contents_regist_link)
    time.sleep(1)
except Exception as e:
    print("Error: コンテンツ登録/更新 メニューが見つかりません。", e)
    driver.quit()
    exit(1)

# ==============================================
# center iframe へ切り替えて検索処理
# ==============================================
driver.switch_to.default_content()
try:
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "center")))
except Exception as e:
    print("Error: center iframe が見つかりません。", e)
    driver.quit()
    exit(1)

try:
    driver.find_element(By.XPATH, "/html/body/div/form/table[1]/tbody/tr[1]/td[1]/input[3]").send_keys(content_id)
    print("検索条件入力完了.")
except:
    print("Error: フォーム入力に失敗しました.")
    exit(1)

try:
    search_button = driver.find_element(By.XPATH, "/html/body/div/form/table[2]/tbody/tr/td/input[1]")
    search_button.click()
    time.sleep(3)
except:
    print("Error: 検索ボタンが見つかりません。")
    exit(1)

# ==============================================
# 一覧から詳細画面へ遷移
# ※iframe内の検索結果一覧にある詳細ボタンをクリックして新ウィンドウを開く
# ==============================================
loopcounter = 5
iframe = driver.find_element(By.XPATH, "/html/body/div/form/iframe")
driver.switch_to.frame(iframe)
try:
    detail_button = driver.find_element(By.XPATH, '//*[@id="SearchResultForm"]/table[1]/tbody/tr[2]/td[6]/input[1]')
    detail_button.click()
    time.sleep(3)
except:
    print("Error: 詳細ボタンが見つかりません。")
    driver.quit()
    exit(1)

# ==============================================
# 新しいウィンドウに切り替え
# ==============================================
try:
    main_window = driver.current_window_handle
    all_windows = driver.window_handles
    for w in all_windows:
        if w != main_window:
            driver.switch_to.window(w)
            break
    time.sleep(2)
    print("新しいウィンドウが表示されました。")
except Exception as e:
    print("Error: 新しいウィンドウへの切り替えに失敗しました。", e)
    driver.quit()
    exit(1)

# ==============================================
# 新ウィンドウで最下部のボタンをクリックし、OKボタンを押下する
# ==============================================
try:
    wait = WebDriverWait(driver, 10)
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(1)
    bottom_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div[5]/form/input[29]")))
    bottom_button.click()
    print("最下部のボタンをクリックしました。")
    time.sleep(2)
    
    try:
        alert = WebDriverWait(driver, 10).until(EC.alert_is_present())
        alert.accept()
        print("アラートのOKボタンを押下しました。")
    except Exception as alert_exception:
        ok_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='OK']")))
        ok_button.click()
        print("ページ上のOKボタンをクリックしました。")
    time.sleep(2)
    
except Exception as e:
    print("Error: 新ウィンドウでのOKボタン押下処理に失敗しました。", e)
    driver.quit()
    exit(1)

# ==============================================
# タイトル枠の値をチェックし、条件に応じた操作を実施
# ==============================================
try:
    # タイトル枠の値を取得
    title_input = driver.find_element(By.XPATH, "/html/body/div/div[5]/form/table[2]/tbody/tr[7]/td/input")
    title_text = title_input.get_attribute("value")
    print("タイトル枠の値:", title_text)
    
    if title_text.startswith("▼"):
        # ───────────────────────────────
        # 【申請】（最下部1番目のボタン）
        application_button = driver.find_element(By.XPATH, "/html/body/div/div[5]/form/table[14]/tbody/tr[3]/td/table/tbody/tr[1]/td/input[1]")
        application_button.click()
        print("【申請】ボタンをクリックしました。")
        time.sleep(2)
        try:
            alert_app = WebDriverWait(driver, 10).until(EC.alert_is_present())
            alert_app.accept()
            print("【申請】確認アラートのOKを押下しました。")
        except Exception as e:
            ok_button_app = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//input[@value='OK']"))
            )
            ok_button_app.click()
            print("【申請】確認画面のページ上のOKボタンをクリックしました。")
        time.sleep(2)
        
        # ───────────────────────────────
        # 【承認】（最下部2番目のボタン）
        approval_button = driver.find_element(By.XPATH, "/html/body/div/div[5]/form/table[14]/tbody/tr[3]/td/table/tbody/tr[1]/td/input[2]")
        approval_button.click()
        print("【承認】ボタンをクリックしました。")
        time.sleep(2)
        try:
            alert_appr = WebDriverWait(driver, 10).until(EC.alert_is_present())
            alert_appr.accept()
            print("【承認】確認アラートのOKを押下しました。")
        except Exception as e:
            ok_button_appr = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//input[@value='OK']"))
            )
            ok_button_appr.click()
            print("【承認】確認画面のページ上のOKボタンをクリックしました。")
        time.sleep(2)
        
        # ───────────────────────────────
        # 【公開】（最下部3番目のボタン ※指定のXPATH）
        publish_button = driver.find_element(By.XPATH, "/html/body/div/div[5]/form/table[14]/tbody/tr[3]/td/table/tbody/tr[1]/td/input[5]")
        publish_button.click()
        print("【公開】ボタンをクリックしました。")
        time.sleep(2)
        try:
            alert_pub = WebDriverWait(driver, 10).until(EC.alert_is_present())
            alert_pub.accept()
            print("【公開】確認アラートのOKを押下しました。")
        except Exception as e:
            ok_button_pub = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//input[@value='OK']"))
            )
            ok_button_pub.click()
            print("【公開】確認画面のページ上のOKボタンをクリックしました。")
        time.sleep(2)
        
    else:
        print("タイトル枠の値は▼で始まっていません。詳細ウィンドウを閉じ、検索画面に戻ります。")
    
    # ウィンドウを閉じて、検索画面に戻る
    driver.close()
    print("詳細ウィンドウを閉じました。")
    
    remaining_windows = driver.window_handles
    if remaining_windows:
        driver.switch_to.window(remaining_windows[0])
        print("検索画面に戻りました。")
    
except Exception as e:
    print("Error: タイトル枠のチェックまたは操作に失敗しました。", e)
    driver.quit()
    exit(1)
