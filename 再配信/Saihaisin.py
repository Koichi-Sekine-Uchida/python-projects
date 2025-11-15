import os
import sys
import time
import win32com.client as win32  # ← 必要なら残してOK。ただし今回は Excel 不要なのでコメントアウト可
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.alert import Alert
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service

# --------------------------------------------------------------------------------
# 【ポイント】
#  1. 「エラー再送ファイル.txt」から9行の情報を取得
#  2. 行の順番:
#       1) アクセスするリンク (URL)
#       2) EduMallのID
#       3) EduMallのPW
#       4) ステータス（配信準備エラー/配信エラー）
#       5) 更新日時_開始 (日付扱いされないよう文字列のまま)
#       6) 更新日時_終了 (日付扱いされないよう文字列のまま & 実行日より前を推奨)
#       7) コンテンツID(任意)
#       8) グループID(任意)
#       9) エッジサーバ名(任意)
# --------------------------------------------------------------------------------

# === 1) テキストファイルのパス設定 ===============================================
current_dir = os.path.dirname(os.path.abspath(__file__))
txt_filename = "エラー再送ファイル.txt"
txt_path = os.path.join(current_dir, txt_filename)

# === 2) テキストファイルから9行を読み込む =======================================
try:
    with open(txt_path, "r", encoding="utf-8") as f:
        lines = f.read().splitlines()
except Exception as e:
    print(f"テキストファイルを開けませんでした: {e}")
    sys.exit()

if len(lines) < 9:
    print("テキストファイルの行数が不足しています。9行必要です。")
    sys.exit()

# 変数に格納
sellSide_url = lines[0].strip()   # アクセスするリンク
edumall_id   = lines[1].strip()   # EduMallのID
edumall_pw   = lines[2].strip()   # EduMallのPW
input_status = lines[3].strip()   # ステータス（配信準備エラー/配信エラー）
start_date   = lines[4].strip()   # 更新日時_開始
end_date     = lines[5].strip()   # 更新日時_終了
content_id   = lines[6].strip()   # コンテンツID(任意)
group_id     = lines[7].strip()   # グループID(任意)
server_name  = lines[8].strip()   # エッジサーバ名(任意)

# === 3) Edge WebDriver のセットアップ ============================================
driver_path = EdgeChromiumDriverManager().install()
service = Service(driver_path)
driver = webdriver.Edge(service=service)
driver.maximize_window()

# === 4) ログイン処理 ===========================================================
driver.get(sellSide_url)
time.sleep(5)

driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[1]/input").send_keys(edumall_id)
driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[2]/input").send_keys(edumall_pw)
driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[3]/button").click()
time.sleep(5)

# CAFS15003 画面へ遷移
driver.get("https://school.edumall.jp/dlvr/CAFS15003")

# === 5) 詳細検索ボタンをクリック ================================================
detail_search_button = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.XPATH, "//label[contains(text(), '詳細検索')]"))
)
driver.execute_script("arguments[0].click();", detail_search_button)
time.sleep(2)

# === 6) グループID 入力（任意） ================================================
if group_id:
    group_id_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//input[@data-bind='value: groupId']"))
    )
    group_id_input.clear()
    group_id_input.send_keys(group_id)

# === 7) エッジサーバ名 入力（任意） ============================================
if server_name:
    server_name_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//input[@data-bind='value: edgeServerName']"))
    )
    server_name_input.clear()
    server_name_input.send_keys(server_name)
    # driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", server_name_input)

# === 8) 更新日時(開始) 入力 ====================================================
print(f"更新日時(開始): '{start_date}' を入力中...")
try:
    start_date_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//input[@data-bind='value: updateDatetimeFrom']"))
    )
    start_date_input.clear()
    time.sleep(1)
    start_date_input.send_keys(start_date)
    driver.execute_script("""
        arguments[0].value = arguments[1];
        arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
    """, start_date_input, start_date)
    time.sleep(1)
    print("更新日時(開始)の入力完了！")
except Exception as e:
    print(f"更新日時(開始)入力中にエラー: {e}")
    sys.exit()

# === 9) 更新日時(終了) 入力 ====================================================
print(f"更新日時(終了): '{end_date}' を入力中...")
try:
    end_date_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//input[@data-bind='value: updateDatetimeTo']"))
    )
    end_date_input.clear()
    time.sleep(1)
    end_date_input.send_keys(end_date)
    driver.execute_script("""
        arguments[0].value = arguments[1];
        arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
    """, end_date_input, end_date)
    time.sleep(1)
    print("更新日時(終了)の入力完了！")
except Exception as e:
    print(f"更新日時(終了)入力中にエラー: {e}")
    sys.exit()

# === 10) 配信エラー以外のチェックボックスを外す =================================
checkbox_xpaths = [
    "/html/body/div[1]/section/div/div/div[3]/div/form/div/div[1]/div[4]/div/label[1]/input",  # 仮予約
    "/html/body/div[1]/section/div/div/div[3]/div/form/div/div[1]/div[4]/div/label[2]/input",  # 配信待ち
    "/html/body/div[1]/section/div/div/div[3]/div/form/div/div[1]/div[4]/div/label[3]/input",  # 配信準備エラー
    "/html/body/div[1]/section/div/div/div[3]/div/form/div/div[1]/div[4]/div/label[4]/input",  # 配信中
    # "/html/body/div[1]/section/div/div/div[3]/div/form/div/div[1]/div[4]/div/label[5]/input",  # 配信エラー (外さない)
    "/html/body/div[1]/section/div/div/div[3]/div/form/div/div[1]/div[4]/div/label[6]/input",  # 配信済
    "/html/body/div[1]/section/div/div/div[3]/div/form/div/div[1]/div[4]/div/label[7]/input",  # 再配信設定済
]

for xpath in checkbox_xpaths:
    checkbox = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, xpath))
    )
    if checkbox.is_selected():
        checkbox.click()

time.sleep(2)

# === 11) 検索ボタンをクリック =================================================
search_button = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//*[@id='mainContens']/div[3]/div/form/div/div[3]/div[3]/div[4]/button"))
)
driver.execute_script("arguments[0].removeAttribute('disabled');", search_button)
driver.execute_script("arguments[0].click();", search_button)
time.sleep(2)

# === 12) 配信エラーを再送する処理 =============================================
flag = True
check_Text = "ディスク容量"  # この単語を含んだら再送しない
saisou_count = 1  # 容量不足のものをスキップするため

while flag:
    time.sleep(1)
    try:
        # テーブルの tr[...] の td[1] をクリック
        row_xpath = f"/html/body/div[1]/section/div/div/div[3]/div/haishin_datatable/div[2]/div/table/tbody/tr[{saisou_count}]/td[1]"
        driver.find_element(By.XPATH, row_xpath).click()
        time.sleep(3)

        # 再配信チェックボックスをクリック
        driver.find_element(By.XPATH, "/html/body/div[1]/section/div/div/div[7]/div[2]/div/div[1]/div/label/span").click()
        time.sleep(1)

        # 「再配信情報更新」ボタンをクリック
        driver.find_element(By.XPATH, "/html/body/div[1]/section/div/div/div[8]/div[2]/button").click()
        time.sleep(1)

        # 配信情報更新ダイアログのOK
        driver.find_element(By.XPATH, "/html/body/div[4]/div/div/div[3]/div/div[2]/button").click()
        time.sleep(2)

        # 2回目のアラート（容量不足や状態不明が出るかどうか）
        try:
            text_element = driver.find_element(By.XPATH, "/html/body/div[4]/div/div/div[2]/p")
            if check_Text in text_element.text:
                # 容量不足なら送信しない
                driver.find_element(By.XPATH, "/html/body/div[4]/div/div/div[3]/div/div[1]/button").click()
                driver.find_element(By.XPATH, "/html/body/div[1]/section/div/div/div[8]/div[1]/a").click()
                saisou_count += 1
            else:
                # 容量不足でないなら送信
                driver.find_element(By.XPATH, "/html/body/div[4]/div/div/div[3]/div/div[2]/button").click()
        except:
            # 追加のポップアップが出なかったらスルー
            pass

        time.sleep(3)
    except Exception as e:
        print("テーブル要素がなくなったか、処理が完了した可能性があります。")
        print(f"エラー内容: {e}")
        flag = False  # ループ終了

print("再送処理が完了しました。")
