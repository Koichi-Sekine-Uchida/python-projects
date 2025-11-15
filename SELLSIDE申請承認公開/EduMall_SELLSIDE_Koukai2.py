import sys
import pythoncom
import win32com.client
import os
import time
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime

# ==============================================
# ログの設定（FileHandler と StreamHandler を併用）
# ==============================================
log_filename = datetime.now().strftime(r"C:\tools\python-projects\SELLSIDE申請承認公開\logs\SELLSIDE申請公開_%Y%m%d_%H%M%S.log")
logger = logging.getLogger("my_logger")
logger.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s %(message)s', '%Y-%m-%d %H:%M:%S')

# ファイルハンドラの設定
file_handler = logging.FileHandler(log_filename, encoding="utf-8")
file_handler.setLevel(logging.INFO)
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)

# コンソールハンドラの設定（PowerShell へ出力）
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setLevel(logging.INFO)
console_handler.setFormatter(formatter)
logger.addHandler(console_handler)

logger.info("スクリプト開始")

# ==============================================
# Excelファイルから必要情報を取得（コンテンツIDはE列、2行目以降）
# ==============================================
filename = r"C:\tools\python-projects\SELLSIDE申請承認公開\配信設定ファイル.xlsx"
if not os.path.exists(filename):
    logger.error(f"エラー: 指定されたExcelファイルが見つかりません: {filename}")
    sys.exit(1)

pythoncom.CoInitialize()
excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False
workbook = excel.Workbooks.Open(filename)
sheet = workbook.Worksheets("Sheet1")

sellSide_url = str(sheet.Cells(1, 2).Value)
edumall_id   = str(sheet.Cells(2, 2).Value)
edumall_pw   = str(sheet.Cells(3, 2).Value)
sleep_time   = int(sheet.Cells(4, 2).Value)

content_ids = []
row = 2
while True:
    cell_val = sheet.Cells(row, 5).Value
    if cell_val is None or str(cell_val).strip() == "":
        break
    content_ids.append(str(cell_val))
    row += 1

workbook.Close(SaveChanges=False)
excel.Quit()

logger.info("取得したコンテンツID: " + str(content_ids))

# ==============================================
# Selenium（Edge）のセットアップ
# ==============================================
driver_path = EdgeChromiumDriverManager().install()
service = Service(driver_path)
driver = webdriver.Edge(service=service)

# ==============================================
# Webアプリへアクセス & ログイン
# ==============================================
try:
    driver.get(sellSide_url)
    time.sleep(3)
    driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[1]/input").send_keys(edumall_id + Keys.TAB)
    driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[2]/input").send_keys(edumall_pw)
    driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[3]/button").click()
    logger.info("ログイン成功")
except Exception as e:
    logger.error("ログイン処理中にエラーが発生しました: " + str(e))
    driver.quit()
    sys.exit(1)

# ==============================================
# メニュー操作（menu iframe へ切り替え、コンテンツ管理→コンテンツ登録/更新）
# ==============================================
try:
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "menu")))
    driver.switch_to.frame("menu")
    contents_menu = driver.find_element(By.XPATH, '//p[@onclick="openMenu(\'1\')"]')
    driver.execute_script("arguments[0].click();", contents_menu)
    time.sleep(1)
    contents_regist_link = driver.find_element(By.XPATH, '//a[@onclick="showPage(this, \'goods/CGdGoodsSearch.jsp\'); return false;"]')
    driver.execute_script("arguments[0].click();", contents_regist_link)
    time.sleep(1)
except Exception as e:
    logger.error("メニュー操作中にエラーが発生しました: " + str(e))
    driver.quit()
    sys.exit(1)

# ==============================================
# 検索画面：center iframe へ切り替え
# ==============================================
driver.switch_to.default_content()
try:
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "center")))
except Exception as e:
    logger.error("Error: center iframe が見つかりません。 " + str(e))
    driver.quit()
    sys.exit(1)

# ==============================================
# 各コンテンツIDでループ処理
# ==============================================
for content_id in content_ids:
    try:
        # ① 検索条件入力
        search_input = driver.find_element(By.XPATH, "/html/body/div/form/table[1]/tbody/tr[1]/td[1]/input[3]")
        search_input.clear()
        search_input.send_keys(content_id)
        logger.info(f"検索条件入力完了: {content_id}")
        
        # ② 検索ボタンを押下
        search_button = driver.find_element(By.XPATH, "/html/body/div/form/table[2]/tbody/tr/td/input[1]")
        search_button.click()
        time.sleep(3)
        
        # ③ 検索結果のiframe へ切り替え
        results_iframe = driver.find_element(By.XPATH, "/html/body/div/form/iframe")
        driver.switch_to.frame(results_iframe)
        
        # ④ 詳細ボタンをクリック
        detail_button = driver.find_element(By.XPATH, '//*[@id="SearchResultForm"]/table[1]/tbody/tr[2]/td[6]/input[1]')
        detail_button.click()
        time.sleep(3)
        
        # ⑤ サブウィンドウ（詳細ウィンドウ）へ切り替え
        main_window = driver.current_window_handle
        all_windows = driver.window_handles
        sub_window = None
        for w in all_windows:
            if w != main_window:
                sub_window = w
                driver.switch_to.window(w)
                break
        time.sleep(2)
        logger.info("詳細ウィンドウが表示されました。")
        
        # ===============================
        # 詳細ウィンドウ内の操作開始（コンテンツ情報の修正）
        # ===============================
        wait = WebDriverWait(driver, 10)
        try:
            # 画面下部までスクロールしてボタンを表示
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1)
            # 「コンテンツ情報の修正」ボタンをクリック
            common_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div[5]/form/input[29]")))
            common_button.click()
            logger.info("コンテンツ情報の修正ボタンをクリックしました。")
            time.sleep(2)
            # アラートが表示された場合の処理
            try:
                alert = wait.until(EC.alert_is_present())
                alert.accept()
                logger.info("コンテンツ情報修正後のアラートOKを押下しました。")
            except Exception:
                ok_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='OK']")))
                ok_button.click()
                logger.info("コンテンツ情報修正後のページ上のOKボタンをクリックしました。")
            time.sleep(2)
        except Exception as e:
            logger.error("Error: コンテンツ情報の修正ボタン処理に失敗しました。 " + str(e))
            # 致命的でなければ次の処理へ進む

        # ───────────────────────────────
        # 【申請】ステージ例
        try:
            application_button = driver.find_element(By.XPATH, "/html/body/div/div[5]/form/table[14]/tbody/tr[3]/td/table/tbody/tr[1]/td/input[1]")
            application_button.click()
            logger.info("【申請】ボタンをクリックしました。")
            time.sleep(2)
            try:
                alert_app = WebDriverWait(driver, 10).until(EC.alert_is_present())
                alert_app.accept()
                logger.info("【申請】確認アラートのOKを押下しました。")
            except Exception:
                ok_button_app = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[@value='OK']")))
                ok_button_app.click()
                logger.info("【申請】確認画面のOKをクリックしました。")
            time.sleep(2)
        except Exception as e:
            logger.error("Error: 【申請】ボタンの押下に失敗しました。 次の【承認】ステージに進みます。 " + str(e))
        
        # ───────────────────────────────
        # 【承認】ステージ例
        try:
            approval_button = driver.find_element(By.XPATH, "/html/body/div/div[5]/form/table[14]/tbody/tr[3]/td/table/tbody/tr[1]/td/input[2]")
            approval_button.click()
            logger.info("【承認】ボタンをクリックしました。")
            time.sleep(2)
            try:
                alert_appr = WebDriverWait(driver, 10).until(EC.alert_is_present())
                alert_appr.accept()
                logger.info("【承認】確認アラートのOKを押下しました。")
            except Exception:
                ok_button_appr = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[@value='OK']")))
                ok_button_appr.click()
                logger.info("【承認】確認画面のOKをクリックしました。")
            time.sleep(2)
        except Exception as e:
            logger.error("Error: 【承認】ボタンの押下に失敗しました。 次の【公開】ステージに進みます。 " + str(e))
        
        # ───────────────────────────────
        # 【公開】ステージ例
        try:
            publish_button = driver.find_element(By.XPATH, "/html/body/div/div[5]/form/table[14]/tbody/tr[3]/td/table/tbody/tr[1]/td/input[5]")
            publish_button.click()
            logger.info("【公開】ボタンをクリックしました。")
            time.sleep(2)
            try:
                alert_pub = WebDriverWait(driver, 10).until(EC.alert_is_present())
                alert_pub.accept()
                logger.info("【公開】確認アラートのOKを押下しました。")
            except Exception:
                ok_button_pub = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[@value='OK']")))
                ok_button_pub.click()
                logger.info("【公開】確認画面のOKをクリックしました。")
            time.sleep(2)
        except Exception as e:
            logger.error("Error: 【公開】ボタンの押下に失敗しました。 サブウィンドウを閉じ、検索画面に戻ります。 " + str(e))
        
        # ───────────────────────────────
        # 詳細ウィンドウ内の操作終了後、サブウィンドウを閉じる
        driver.close()
        logger.info("詳細ウィンドウを閉じました。")
        
        # メインウィンドウ（検索画面）に戻る
        driver.switch_to.window(main_window)
        logger.info("検索画面に戻りました。")
        
        # 検索画面の center iframe へ再度切り替え
        driver.switch_to.default_content()
        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "center")))
        time.sleep(2)
        
        logger.info(f"処理完了: {content_id}")
        time.sleep(2)
    
    except Exception as loop_err:
        logger.error(f"Error processing content_id {content_id}: " + str(loop_err))
        try:
            if len(driver.window_handles) > 1:
                driver.switch_to.window(driver.window_handles[-1])
                driver.close()
            driver.switch_to.window(main_window)
            driver.switch_to.default_content()
            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "center")))
        except:
            pass
    
    logger.info(f"完了: {content_id}")
    time.sleep(2)

logger.info("全ての検索処理が完了しました。")
input("Enterキーを押すと終了します...")
driver.quit()
