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

#20250129 END


#設定ファイル読込
filename = 'C:/tools/selenium/再配信/エラー再送ファイル2.xlsx'
#Excelを開く
xlApps = win32.Dispatch("Excel.Application")
workbook  = xlApps.Workbooks.Open(filename)
sheet = workbook.Worksheets("Sheet1")
#EduMall_ID
edumall_id = str(sheet.Cells(2,2))
#EduMall_PW
edumall_pw = str(sheet.Cells(3,2))
#Sell-SideのURL
sellSide_url = str(sheet.Cells(1,2))

input_status = str(sheet.Cells(4,2))

start_year = str(sheet.Cells(5,2))
end_year = str(sheet.Cells(6,2))
content_id = str(sheet.Cells(7,2))
group_id = str(sheet.Cells(8,2))




#20250129 START
#from selenium.webdriver.support.ui import WebDriverWait
#from selenium.webdriver.support import expected_conditions as EC

#WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/header/nav/div[1]/a/em"))).click()
#20250129 END



#アクセスするリンク
driver.get(sellSide_url)
time.sleep(5)
driver.find_element(By.XPATH,"/html/body/div/div[2]/form/div[1]/input").send_keys(edumall_id)
driver.find_element(By.XPATH,"/html/body/div/div[2]/form/div[2]/input").send_keys(edumall_pw)
driver.find_element(By.XPATH,"/html/body/div/div[2]/form/div[3]/button").click()

#20250129 START
driver.get("https://school.edumall.jp/dlvr/CAFS15003")
time.sleep(5)
#20250129 END


from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

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

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# 「詳細検索」ボタンをクリック（すでに展開されていない場合）
detail_search_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//label[contains(text(), '詳細検索')]"))
)
detail_search_button.click()
time.sleep(2)  # 展開が完了するのを待つ

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

# 「グループID」を入力
if sheet.Cells(8,2).value is not None:
    group_id_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//input[@data-bind='value: groupId']"))
    )
    group_id_input.clear()
    group_id_input.send_keys(group_id)

# 検索ボタンを取得し、JavaScript でクリック
search_button = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//button[contains(text(), '検索')]"))
)

# JavaScript でボタンを有効化
driver.execute_script("arguments[0].removeAttribute('disabled');", search_button)

# JavaScript でクリックを実行
driver.execute_script("arguments[0].click();", search_button)
'''

# JavaScript でボタンを有効化
driver.execute_script("arguments[0].removeAttribute('disabled');", search_button)

# JavaScript でクリックを実行
driver.execute_script("arguments[0].click();", search_button)

'''
'''
# 検索ボタンを探してクリック
search_button = WebDriverWait(driver, 10).until(
#    EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/section/div/div/div[3]/div/form/div/div[3]/div[3]/div[4]/button"))
    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), '検索')]"))
)
search_button.click()
'''




'''
#配信情報参照
driver.find_element(By.XPATH,"/html/body/div[1]/header/nav/div[1]/a/em").click()
driver.find_element(By.XPATH,"/html/body/div[1]/aside/div/nav/ul/li[4]/a/em").click()
time.sleep(1)
driver.find_element(By.XPATH,"/html/body/div[1]/aside/div/nav/ul/li[4]/ul/li[3]/a").click()
time.sleep(10)
#配信準備エラーを選択
driver.find_element(By.XPATH,"/html/body/div[1]/section/div/div/div[3]/div/form/div/div[1]/div[4]/div/label[1]/input").click() #[仮予約]のチェックボックスを外す
driver.find_element(By.XPATH,"/html/body/div[1]/section/div/div/div[3]/div/form/div/div[1]/div[4]/div/label[2]/input").click() #[配信待ち]のチェックボックスを外す
if input_status == '配信エラー':
    driver.find_element(By.XPATH,"/html/body/div[1]/section/div/div/div[3]/div/form/div/div[1]/div[4]/div/label[3]/input").click() #ステータスが[配信エラー]の場合、[配信準備エラー]のチェックボックスを外す
driver.find_element(By.XPATH,"/html/body/div[1]/section/div/div/div[3]/div/form/div/div[1]/div[4]/div/label[4]/input").click() #[配信中]のチェックボックスを外す
if input_status == '配信準備エラー':
    driver.find_element(By.XPATH,"/html/body/div[1]/section/div/div/div[3]/div/form/div/div[1]/div[4]/div/label[5]/input").click() #ステータスが[配信準備エラー]の場合、[配信エラー]のチェックボックスを外す
driver.find_element(By.XPATH,"/html/body/div[1]/section/div/div/div[3]/div/form/div/div[1]/div[4]/div/label[6]/input").click() #[配信済]のチェックボックスを外す
driver.find_element(By.XPATH,"/html/body/div[1]/section/div/div/div[3]/div/form/div/div[1]/div[4]/div/label[7]/input").click() #[再配信設定済]のチェックボックスを外す
'''
'''
dropdown = driver.find_element(By.XPATH,"/html/body/div[1]/section/div/div/div[3]/div/form/div/div[1]/div[4]/select")
select = Select(dropdown)
select.select_by_visible_text(input_status)
'''
'''
#driver.find_element(By.XPATH,"/html/body/div[1]/section/div/div/div[3]/div/form/div/div[2]/div[2]/label/em").click()
#driver.find_element(By.XPATH,"/html/body/div[1]/section/div/div/div[3]/div/form/div/div[3]/div[2]/div[2]/div/div[1]/div/input").send_keys(start_year)
#driver.find_element(By.XPATH,"/html/body/div[1]/section/div/div/div[3]/div/form/div/div[3]/div[2]/div[2]/div/div[3]/div/input").send_keys(end_year)

#if(sheet.Cells(7,2).value is not None):
#    driver.find_element(By.XPATH,"/html/body/div[1]/section/div/div/div[3]/div/form/div/div[3]/div[3]/div[1]/input").send_keys(content_id)
#if(sheet.Cells(8,2).value is not None):
    driver.find_element(By.XPATH,"/html/body/div[1]/section/div/div/div[3]/div/form/div/div[1]/div[2]/input").send_keys(group_id)
#driver.find_element(By.XPATH,"/html/body/div[1]/section/div/div/div[3]/div/form/div/div[3]/div[3]/div[4]/button").click()
#time.sleep(3)
dropdown = driver.find_element(By.XPATH,"/html/body/div[1]/section/div/div/div[3]/div/haishin_datatable/div[1]/div/div/div/div/select")
select = Select(dropdown)
select.select_by_visible_text("100")
time.sleep(3)
'''

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