from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# ウェブドライバーの設定
driver = webdriver.Chrome() # または別のブラウザを指定

# ExcelファイルからグループIDを読み込む
group_ids = data['グループID'].unique()

for group_id in group_ids:
    # ウェブサイトにアクセス
    driver.get("https://school.edumall.jp/dlvr/CAFS15003")
    
    # グループIDを入力欄に入力
    group_id_input = driver.find_element(By.ID, 'group-id-input') # 実際の要素IDに置き換える
    group_id_input.clear()
    group_id_input.send_keys(group_id)
    
    # 検索ボタンをクリック
    search_button = driver.find_element(By.ID, 'search-button') # 実際の要素IDに置き換える
    search_button.click()
    
    # 結果がロードされるのを待つ
    time.sleep(2) # 必要に応じて調整

    # 配信ステータスを確認
    try:
        # 配信ステータスが「配信済」である行を探す (この部分は実際のウェブサイトに応じて適宜調整)
        delivered_status = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//tr[td[text()='【配信済】']]"))
        )
        
        # 配信IDをクリック
        delivery_id = delivered_status.find_element(By.XPATH, ".//td[1]") # 実際の要素に置き換える
        delivery_id.click()
        
        # 再配信チェックボタンにチェックを入れる
        resend_checkbox = driver.find_element(By.ID, 'resend-checkbox') # 実際の要素IDに置き換える
        if not resend_checkbox.is_selected():
            resend_checkbox.click()
        
        # 更新ボタンをクリック
        update_button = driver.find_element(By.ID, 'update-button') # 実際の要素IDに置き換える
        update_button.click()
        
    except TimeoutException:
        print(f"Group ID {group_id} has no delivered status or other issue occurred.")
        continue

# ブラウザを閉じる
driver.quit()
