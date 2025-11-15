import win32com.client as win32
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.alert import Alert
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


#ドライバーの自動インストール
#driver = webdriver.Edge(EdgeChromiumDriverManager().install())
#20250129 START
from selenium.webdriver.edge.service import Service

# ドライバのパスを取得
driver_path = EdgeChromiumDriverManager().install()

# サービスオブジェクトを作成
service = Service(driver_path)

# WebDriverのインスタンスを作成
driver = webdriver.Edge(service=service)
driver.maximize_window()



#設定ファイル読込
filename = r"C:\tools\python-projects\再配信\エラー再送ファイル.xlsx"
# 早期バインディングを利用して Excel アプリケーションを取得
xlApps = win32.gencache.EnsureDispatch("Excel.Application")

# Excel ファイルを開く
workbook = xlApps.Workbooks.Open(filename)
sheet = workbook.Worksheets("Sheet1")

#EduMall_ID
edumall_id = str(sheet.Cells(2,2))
#EduMall_PW
edumall_pw = str(sheet.Cells(3,2))
#Sell-SideのURL
sellSide_url = str(sheet.Cells(1,2))

input_status = str(sheet.Cells(4,2))

start_date = str(sheet.Cells(5,2))
end_date = str(sheet.Cells(6,2))
content_id = str(sheet.Cells(7,2))
group_id = str(sheet.Cells(8,2))
server_name = str(sheet.Cells(9,2))

#アクセスするリンク
driver.get(sellSide_url)
time.sleep(5)
driver.find_element(By.XPATH,"/html/body/div/div[2]/form/div[1]/input").send_keys(edumall_id)
driver.find_element(By.XPATH,"/html/body/div/div[2]/form/div[2]/input").send_keys(edumall_pw)
driver.find_element(By.XPATH,"/html/body/div/div[2]/form/div[3]/button").click()
time.sleep(5)

driver.get("https://school.edumall.jp/dlvr/CAFS15003")


# 「詳細検索」ボタンをクリック（すでに展開されていない場合）
detail_search_button = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.XPATH, "//label[contains(text(), '詳細検索')]"))
)
driver.execute_script("arguments[0].click();", detail_search_button)
time.sleep(2)  # 展開が完了するのを待つ


# 「グループID」を入力
if sheet.Cells(8,2).value is not None:
    group_id_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//input[@data-bind='value: groupId']"))
    )
    group_id_input.clear()
    group_id_input.send_keys(group_id)

# 「エッジサーバ名」を入力
if sheet.Cells(9,2).value is not None:
    server_name_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//input[@data-bind='value: edgeServerName']"))
    )
    server_name_input.clear()
    server_name_input.send_keys(server_name)
    # イベントを発火して変更を反映させる
#    driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", server_name_input)

# ── 更新日時の入力（配信予定日時欄） ─────────────────────
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
    cleanup_excel()
    sys.exit()
    
print(f"配信予定日時(終了): '{end_date}' を入力中...")
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
    cleanup_excel()
    sys.exit()

'''
# 「更新日時」の左側の入力フィールドに start_year をセット（JavaScript使用）
update_start = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//input[@data-bind='value: updateDatetimeFrom']"))
)
driver.execute_script("arguments[0].value = arguments[1];", update_start, start_year)
driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", update_start)

# 「更新日時」の右側の入力フィールドに end_year をセット（JavaScript使用）
update_end = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//input[@data-bind='value: updateDatetimeTo']"))
)
driver.execute_script("arguments[0].value = arguments[1];", update_end, end_year)
driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", update_end)

# 「コンテンツID」を入力
if sheet.Cells(7,2).value is not None:
    content_id_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//input[@data-bind='value: contentsId']"))
    )
    content_id_input.clear()
    content_id_input.send_keys(content_id)
'''

# 配信エラー以外のチェックボックスの XPATH
checkbox_xpaths = [
    "/html/body/div[1]/section/div/div/div[3]/div/form/div/div[1]/div[4]/div/label[1]/input",  # 仮予約
    "/html/body/div[1]/section/div/div/div[3]/div/form/div/div[1]/div[4]/div/label[2]/input",  # 配信待ち
    "/html/body/div[1]/section/div/div/div[3]/div/form/div/div[1]/div[4]/div/label[3]/input",  # 配信準備エラー
    "/html/body/div[1]/section/div/div/div[3]/div/form/div/div[1]/div[4]/div/label[4]/input",  # 配信中
#    "/html/body/div[1]/section/div/div/div[3]/div/form/div/div[1]/div[4]/div/label[5]/input",  # 配信エラー
    "/html/body/div[1]/section/div/div/div[3]/div/form/div/div[1]/div[4]/div/label[6]/input",  # 配信済
    "/html/body/div[1]/section/div/div/div[3]/div/form/div/div[1]/div[4]/div/label[7]/input",  # 再配信設定済
]

# 各チェックボックスの選択を確認し、チェックが入っていたら外す
for xpath in checkbox_xpaths:
    checkbox = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, xpath))
    )
    if checkbox.is_selected():
        checkbox.click()  # クリックしてチェックを外す


time.sleep(2)  # 展開が完了するのを待つ

# 指定の絶対XPathを使って検索ボタンを取得
search_button = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//*[@id='mainContens']/div[3]/div/form/div/div[3]/div[3]/div[4]/button"))
)

# JavaScriptでボタンのdisabled属性を削除してクリック
driver.execute_script("arguments[0].removeAttribute('disabled');", search_button)
driver.execute_script("arguments[0].click();", search_button)
time.sleep(2)  # 展開が完了するのを待つ


flag = True
check_Text = "ディスク容量"#この単語を含んだら再送しない
saisou_count = 1 #容量不足のものをスキップするため

#処理開始(無限ループとなるが、終わったら選択対象がなくなり止まるので無問題)
while(flag == True):
    time.sleep(1)
    driver.find_element(By.XPATH,"/html/body/div[1]/section/div/div/div[3]/div/haishin_datatable/div[2]/div/table/tbody/tr["+str(saisou_count)+"]/td[1]").click()
    time.sleep(3)
    driver.find_element(By.XPATH,"/html/body/div[1]/section/div/div/div[7]/div[2]/div/div[1]/div/label/span").click()
    time.sleep(1)
    driver.find_element(By.XPATH,"/html/body/div[1]/section/div/div/div[8]/div[2]/button").click()
    time.sleep(1)
    #配信情報更新
    driver.find_element(By.XPATH,"/html/body/div[4]/div/div/div[3]/div/div[2]/button").click()
    time.sleep(2)
    #2回目のアラート（容量不足や状態不明が出るかどうか）
    try:
        text = driver.find_element(By.XPATH,"/html/body/div[4]/div/div/div[2]/p")
        hantei = check_Text in text.text
        if(hantei == False):
            #容量不足でないなら、送信
            driver.find_element(By.XPATH,"/html/body/div[4]/div/div/div[3]/div/div[2]/button").click()
        else:
            #容量不足なら送信しない
            driver.find_element(By.XPATH,"/html/body/div[4]/div/div/div[3]/div/div[1]/button").click()
            driver.find_element(By.XPATH,"/html/body/div[1]/section/div/div/div[8]/div[1]/a").click()
            saisou_count = saisou_count + 1
    except:
        pass
    time.sleep(3)