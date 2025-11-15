import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.service import Service

# ドライバの自動インストール
driver_path = EdgeChromiumDriverManager().install()
service = Service(driver_path)
driver = webdriver.Edge(service=service)

# 設定ファイルから情報を取得（例として手動入力）
sellSide_url = "https://example.com"  # 実際のURLを設定
edumall_id = "your_username"
edumall_pw = "your_password"

# 1. ログインページにアクセス
driver.get(sellSide_url)
time.sleep(5)

# 2. ログイン情報を入力
driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[1]/input").send_keys(edumall_id)
driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[2]/input").send_keys(edumall_pw)
driver.find_element(By.XPATH, "/html/body/div/div[2]/form/div[3]/button").click()
time.sleep(5)

# 3. `menu` の `iframe` に切り替え
WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "menu")))

# 4. `menu-wrap` が表示されるまで待つ
WebDriverWait(driver, 10).until(
    EC.visibility_of_element_located((By.ID, "menu-wrap"))
)

# 5. 「ACCIS注文連携」のリンクをクリック
accis_order_link = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'ACCIS注文連携')]"))
)
accis_order_link.click()
time.sleep(5)

# 6. メインページに戻る
driver.switch_to.default_content()

# 7. `center` の `iframe` に切り替え（必要なら）
WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, "center")))

# 8. `center` 内の要素を操作（例として適当な操作を入れる）
try:
    some_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//h1[contains(text(), 'ACCIS注文連携')]"))
    )
    print("ACCIS注文連携ページが開かれました")
except:
    print("ページ遷移に失敗しました")

# 9. メインページに戻る
driver.switch_to.default_content()

# 処理終了後にドライバを閉じる（テスト時はコメントアウト）
# driver.quit()
