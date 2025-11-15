import sys
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

# ----------------------------------------------
# ログ保存用ディレクトリの作成
# ----------------------------------------------
log_dir = r"C:\tools\python-projects\SELLSIDE実体公開テキスト版\logs"
if not os.path.exists(log_dir):
    os.makedirs(log_dir)

# ----------------------------------------------
# ログファイルのパス定義（指定通りに修正）
# ----------------------------------------------
log_filename = datetime.now().strftime(r"C:\\tools\\python-projects\\SELLSIDE実体公開テキスト版\\logs\\SELLSIDE申請公開_%Y%m%d_%H%M%S.log")
logger = logging.getLogger("my_logger")
logger.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s %(message)s', '%Y-%m-%d %H:%M:%S')

file_handler = logging.FileHandler(log_filename, encoding="utf-8")
file_handler.setLevel(logging.INFO)
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)

console_handler = logging.StreamHandler(sys.stdout)
console_handler.setLevel(logging.INFO)
console_handler.setFormatter(formatter)
logger.addHandler(console_handler)

logger.info("スクリプト開始")

# ----------------------------------------------
# ログイン情報を「ログイン.txt」から取得
# ----------------------------------------------
login_filename = r"C:\tools\python-projects\SELLSIDE実体公開テキスト版\ログイン.txt"
if not os.path.exists(login_filename):
    logger.error(f"エラー: 指定されたログインファイルが見つかりません: {login_filename}")
    sys.exit(1)

with open(login_filename, "r", encoding="utf-8") as f:
    lines = f.read().splitlines()

if len(lines) < 4:
    logger.error("エラー: ログイン情報ファイルの行数が不足しています。")
    sys.exit(1)

sellSide_url = lines[0].strip()
edumall_id   = lines[1].strip()
edumall_pw   = lines[2].strip()
try:
    sleep_time = int(lines[3].strip())
except ValueError:
    logger.error("エラー: sleep_time の値が整数ではありません。")
    sys.exit(1)

# ----------------------------------------------
# コンテンツIDを「リスト.txt」から取得
# ----------------------------------------------
list_filename = r"C:\tools\python-projects\SELLSIDE実体公開テキスト版\リスト.txt"
if not os.path.exists(list_filename):
    logger.error(f"エラー: 指定されたコンテンツIDリストファイルが見つかりません: {list_filename}")
    sys.exit(1)

with open(list_filename, "r", encoding="utf-8") as f:
    content_ids = [line.strip() for line in f if line.strip() != ""]

logger.info("取得したコンテンツID: " + str(content_ids))

# ----------------------------------------------
# Selenium（Edge）のセットアップ
# ----------------------------------------------
driver_path = EdgeChromiumDriverManager().install()
service = Service(driver_path)
driver = webdriver.Edge(service=service)

# ----------------------------------------------
# Webアプリへアクセス & ログイン
# ----------------------------------------------
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

# ----------------------------------------------
# メニュー操作（menu iframe へ切り替え、コンテンツ管理→コンテンツ登録/更新）
# ----------------------------------------------
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

# ----------------------------------------------
# 検索画面：center iframe へ切り替え
# ----------------------------------------------
driver.switch_to.default_content()
try:
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "center")))
except Exception as e:
    logger.error("Error: center iframe が見つかりません。 " + str(e))
    driver.quit()
    sys.exit(1)

# ----------------------------------------------
# 各コンテンツIDでループ処理
# ----------------------------------------------
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
        
        # ★ 最初にメッセージがあるかチェック（押下前）
        try:
            message_elem = driver.find_element(By.XPATH, "/html/body/div/div[5]/form/table[14]/tbody/tr[4]/td/table/tbody/tr[2]/td/div")
            message_text = message_elem.text.strip()
        except Exception:
            message_text = ""

        if message_text:
            logger.info("初期メッセージが存在します: " + message_text)
            # 追加：ステータスチェック
            try:
                final_status = driver.find_element(By.XPATH, "/html/body/div/div[5]/form/table[14]/tbody/tr[1]/td[2]").text.strip()
            except Exception as e:
                final_status = ""
                logger.error("ステータス取得エラー: " + str(e))
            if final_status == "承認（公開前）":
                with open(r"C:\tools\python-projects\SELLSIDE実体公開テキスト版\公開前リスト.txt", "a", encoding="utf-8") as f_out:
                    f_out.write(content_id + "\n")
                logger.info(f"{content_id} は『承認（公開前）』の状態のため抽出されました。")
            
            driver.close()
            driver.switch_to.window(main_window)
            driver.switch_to.default_content()
            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "center")))
            logger.info(f"{content_id} の処理をスキップします。")
            continue

        # ステータス取得（ボタン押下前）
        try:
            status_text = driver.find_element(By.XPATH, "/html/body/div/div[5]/form/table[14]/tbody/tr[1]/td[2]").text.strip()
            logger.info("初期ステータス: " + status_text)
        except Exception as e:
            logger.error("ステータスの取得に失敗しました: " + str(e))
            driver.close()
            driver.switch_to.window(main_window)
            driver.switch_to.default_content()
            WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "center")))
            continue
        
        # ★ ステータスが「承認（公開済み）」の「未申請」場合は、コンテンツ実体の修正ボタンを押下する
        if status_text in ["承認（公開済み）", "未申請"]:
            try:
                modify_button = driver.find_element(By.XPATH, "/html/body/div/div[5]/form/input[30]")
                modify_button.click()
                logger.info("コンテンツ実体の修正ボタンをクリックしました。")
                time.sleep(2)
                try:
                    alert = WebDriverWait(driver, 10).until(EC.alert_is_present())
                    alert.accept()
                    logger.info("コンテンツ実体修正後のアラートを承認しました。")
                except Exception:
                    ok_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[@value='OK']")))
                    ok_button.click()
                    logger.info("コンテンツ実体修正後のOKボタンをクリックしました。")
                time.sleep(2)
                # 修正後、再度ステータスを取得
                status_text = driver.find_element(By.XPATH, "/html/body/div/div[5]/form/table[14]/tbody/tr[1]/td[2]").text.strip()
                logger.info("修正後のステータス: " + status_text)
            except Exception as e:
                logger.error("コンテンツ実体の修正ボタン押下に失敗しました: " + str(e))
        
        # ★ 各ボタン（申請／承認／公開）の要素を取得
        try:
            application_button = driver.find_element(By.XPATH, "//*[@id='InputFields13']/tbody/tr[4]/td/table/tbody/tr[1]/td/input[1]")
            approval_button    = driver.find_element(By.XPATH, "//*[@id='InputFields13']/tbody/tr[4]/td/table/tbody/tr[1]/td/input[2]")
            publish_button     = driver.find_element(By.XPATH, "//*[@id='InputFields13']/tbody/tr[4]/td/table/tbody/tr[1]/td/input[5]")
        except Exception as e:
            logger.error("Error: ボタン要素の取得に失敗しました: " + str(e))
            continue

        # ① 「申請」ボタンが有効なら押下
        if application_button.is_enabled():
            try:
                application_button.click()
                logger.info("実体【申請】ボタンをクリックしました。")
                time.sleep(2)
                try:
                    alert_app = WebDriverWait(driver, 10).until(EC.alert_is_present())
                    alert_app.accept()
                    logger.info("実体【申請】確認アラートのOKを押下しました。")
                except Exception:
                    ok_button_app = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//input[@value='OK']"))
                    )
                    ok_button_app.click()
                    logger.info("実体【申請】確認画面のOKをクリックしました。")
                time.sleep(2)
                # 申請押下後、ステータス再取得
                status_text = driver.find_element(By.XPATH, "/html/body/div/div[5]/form/table[14]/tbody/tr[1]/td[2]").text.strip()
                logger.info("申請押下後のステータス: " + status_text)
            except Exception as e:
                logger.error("Error: 実体【申請】ボタンの押下に失敗しました。 " + str(e))
        else:
            logger.info("【申請】ボタンは無効のためスキップします。")

        # ② 「承認」ボタンが有効なら押下
        if approval_button.is_enabled():
            try:
                approval_button.click()
                logger.info("実体【承認】ボタンをクリックしました。")
                time.sleep(2)
                try:
                    alert_appr = WebDriverWait(driver, 10).until(EC.alert_is_present())
                    alert_appr.accept()
                    logger.info("実体【承認】確認アラートのOKを押下しました。")
                except Exception:
                    ok_button_appr = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//input[@value='OK']"))
                    )
                    ok_button_appr.click()
                    logger.info("実体【承認】確認画面のOKをクリックしました。")
                time.sleep(2)
                # 承認押下後、ステータス再取得
                status_text = driver.find_element(By.XPATH, "/html/body/div/div[5]/form/table[14]/tbody/tr[1]/td[2]").text.strip()
                logger.info("承認押下後のステータス: " + status_text)
            except Exception as e:
                logger.error("Error: 実体【承認】ボタンの押下に失敗しました。 " + str(e))
        else:
            logger.info("【承認】ボタンは無効のためスキップします。")

        # ③ 「公開」ボタンについて
        if publish_button.is_enabled():
            try:
                publish_button.click()
                logger.info("【公開】ボタンをクリックしました。")
                time.sleep(2)
                try:
                    alert_pub = WebDriverWait(driver, 10).until(EC.alert_is_present())
                    alert_pub.accept()
                    logger.info("【公開】確認アラートのOKを押下しました。")
                except Exception:
                    ok_button_pub = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//input[@value='OK']"))
                    )
                    ok_button_pub.click()
                    logger.info("【公開】確認画面のOKをクリックしました。")
                time.sleep(2)
                # 公開押下後、ステータス再取得
                status_text = driver.find_element(By.XPATH, "/html/body/div/div[5]/form/table[14]/tbody/tr[1]/td[2]").text.strip()
                logger.info("公開押下後のステータス: " + status_text)
            except Exception as e:
                logger.error("Error: 【公開】ボタンの押下に失敗しました。 " + str(e))
        else:
            # 公開ボタンが無効の場合、かつメッセージが存在すれば最終ステータスチェックを実施
            try:
                updated_message_elem = driver.find_element(By.XPATH, "/html/body/div/div[5]/form/table[14]/tbody/tr[4]/td/table/tbody/tr[2]/td/div")
                updated_message_text = updated_message_elem.text.strip()
                if updated_message_text:
                    logger.info("【公開】ボタンが無効かつメッセージが存在します: " + updated_message_text)
                    try:
                        final_status = driver.find_element(By.XPATH, "/html/body/div/div[5]/form/table[14]/tbody/tr[1]/td[2]").text.strip()
                    except Exception as e:
                        final_status = ""
                    if final_status == "承認（公開前）":
                        with open(r"C:\tools\python-projects\SELLSIDE実体公開テキスト版\公開前リスト.txt", "a", encoding="utf-8") as f_out:
                            f_out.write(content_id + "\n")
                        logger.info(f"{content_id} は『承認（公開前）』の状態のため抽出されました。")
                    # 該当メッセージがあれば以降の処理は中断して次のコンテンツへ
                    driver.close()
                    driver.switch_to.window(main_window)
                    driver.switch_to.default_content()
                    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "center")))
                    continue
            except Exception as e:
                logger.error("Error: 公開ボタンが無効の場合のメッセージ確認に失敗しました。 " + str(e))
                
        # ───────────────────────────────
        # ★ 処理終了後、詳細ウィンドウを閉じ、メインウィンドウに戻る ★
        driver.close()
        logger.info("詳細ウィンドウを閉じました。")
        driver.switch_to.window(main_window)
        logger.info("検索画面に戻りました。")
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
