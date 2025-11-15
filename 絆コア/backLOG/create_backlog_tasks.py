from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from datetime import datetime
import time
import re

# ----------- 設定 -----------
BACKLOG_URL = "https://ucdprj.backlog.com/add/KIZUNACORE_OPERATION"
WAIT_TIME = 10
実績時間 = "1"
# ----------------------------

def setup_driver():
    options = EdgeOptions()
    options.add_argument("--start-maximized")
    service = EdgeService(EdgeChromiumDriverManager().install())
    return webdriver.Edge(service=service, options=options)

def create_issue(driver, summary, description, parent_key=None):
    driver.get(BACKLOG_URL)
    wait = WebDriverWait(driver, WAIT_TIME)

    # 件名
    wait.until(EC.presence_of_element_located((By.ID, "summary"))).send_keys(summary)

    # 説明
    driver.find_element(By.ID, "description").send_keys(description)

    # 状態「完了」
    status_dropdown = wait.until(EC.presence_of_element_located((By.NAME, "statusId")))
    for option in status_dropdown.find_elements(By.TAG_NAME, "option"):
        if option.text.strip() == "完了":
            option.click()
            break

    # 実績時間
    driver.find_element(By.NAME, "actualHours").send_keys(実績時間)

    # 親課題指定（子課題のみ）
    if parent_key:
        parent_input = driver.find_element(By.NAME, "parentIssueKey")
        parent_input.send_keys(parent_key)
        time.sleep(1)
        parent_input.send_keys("\n")

    # 追加ボタン
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'追加')]"))).click()
    print(f"課題登録完了: {summary}")

def get_latest_issue_key(driver):
    wait = WebDriverWait(driver, WAIT_TIME)
    wait.until(EC.url_contains("/view/"))
    current_url = driver.current_url
    match = re.search(r"/view/([A-Z]+-\d+)", current_url)
    return match.group(1) if match else None

def main():
    today = datetime.now().strftime("%Y年%m月%d日")
    親件名 = f"{today} #Core日次監視（親）"
    親説明 = "09時00分に日次監視を実施\n親課題です。"

    子件名 = f"{today} #Core日次監視（予）"
    子説明 = """09時00分 日次監視
☑ 1. FrontDoor（エラー1件以上）
☑ 2. Application Insightsのログ確認
☑ 3. AppServicePlan（100%稼働）
☑ 4. AppService（http 5xx件以上）
☑ 5. FrontDoor（FDヘルスフロー）
☑ 6. AppServicePlan（6個以上のインスタンス）
☑ 7. VPN接続チェック
☑ 8. SQL接続NGリスト"""

    driver = setup_driver()
    driver.get(BACKLOG_URL)
    input("ログイン完了後にEnterを押してください：")

    create_issue(driver, 親件名, 親説明)
    parent_key = get_latest_issue_key(driver)

    if parent_key:
        print(f"親課題キー：{parent_key}")
        create_issue(driver, 子件名, 子説明, parent_key)
    else:
        print("親課題キーの取得に失敗しました。")

    time.sleep(3)
    driver.quit()

if __name__ == "__main__":
    main()
