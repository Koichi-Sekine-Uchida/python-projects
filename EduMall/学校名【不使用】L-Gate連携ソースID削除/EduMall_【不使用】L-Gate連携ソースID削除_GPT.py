# -*- coding: utf-8 -*-
"""
EduMall 学校一覧（学校名=【不使用】）で
「L-Gate連携ソースID」が入っている行だけ詳細に入り、
詳細画面の「L-Gate連携ソースID」を空欄にして更新する。
（設定.xlsx 版 / URL・ID・PW は B1/B2/B3）
"""

import os
import time
import gc
import traceback

import win32com.client as win32
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ================= 設定 =================
SCHOOL_LIST_URL = "https://school.edumall.jp/schl/CAAS11001"
TARGET_SCHOOL_NAME = "【不使用】"
EXPLICIT_WAIT_SEC = 15

# ================= Edge起動 =================
edge_options = EdgeOptions()
edge_options.use_chromium = True
edge_options.add_experimental_option("detach", True)  # 動作確認中はブラウザ保持
driver = webdriver.Edge(options=edge_options)  # Selenium Manager に任せる
wait = WebDriverWait(driver, EXPLICIT_WAIT_SEC)

# ================= Excelから設定取得 =================
def cleanup_excel():
    try:
        workbook.Close(False)
        xlApp.Quit()
    except:
        pass
    finally:
        try:
            del sheet, workbook, xlApp
        except:
            pass
        gc.collect()

def get_excel_value(r, c):
    v = sheet.Cells(r, c).Value
    return "" if v is None else str(v).strip()

current_dir = os.path.dirname(os.path.abspath(__file__))
excel_path = os.path.join(current_dir, "設定.xlsx")
xlApp = win32.Dispatch("Excel.Application")
workbook = xlApp.Workbooks.Open(excel_path)
sheet = workbook.Worksheets("Sheet1")
BASE_URL = get_excel_value(1, 2)  # B1
LOGIN_ID = get_excel_value(2, 2)  # B2
LOGIN_PW = get_excel_value(3, 2)  # B3

# ================= 共通関数 =================
def safe_click(element):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
    time.sleep(0.1)
    driver.execute_script("arguments[0].click();", element)

def fill_and_fire(el, value):
    driver.execute_script("""
        arguments[0].value = arguments[1];
        arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
        arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
    """, el, value)

def find_by_label_input(label_text):
    xps = [
        f"(.//label[contains(normalize-space(.),'{label_text}')])[1]/following::input[1]",
        f"//*[@id='mainContens']//label[contains(.,'{label_text}')]/following::input[1]",
    ]
    for xp in xps:
        try:
            el = wait.until(EC.presence_of_element_located((By.XPATH, xp)))
            if el.is_displayed():
                return el
        except:
            continue
    return None

def get_column_index_by_header_text(header_text):
    ths = driver.find_elements(By.XPATH, "//table/thead//th")
    for idx, th in enumerate(ths, start=1):
        if header_text in th.text.strip():
            return idx
    return None

def get_rows_xpath_base():
    return "//*[@id='mainContens']/div[3]/div/datatable/div[2]/div/table/tbody"

def get_rows():
    return wait.until(EC.presence_of_all_elements_located((By.XPATH, f"{get_rows_xpath_base()}/tr")))

def get_cell_text(row_idx, col_idx):
    xp = f"{get_rows_xpath_base()}/tr[{row_idx}]/td[{col_idx}]"
    return driver.find_element(By.XPATH, xp).text.strip()

def click_row_school_id(row_idx=1):
    for xp in [
        f"{get_rows_xpath_base()}/tr[{row_idx}]/td[1]//span",
        f"{get_rows_xpath_base()}/tr[{row_idx}]/td[1]//a",
    ]:
        try:
            el = wait.until(EC.element_to_be_clickable((By.XPATH, xp)))
            safe_click(el)
            return
        except:
            continue
    raise RuntimeError("学校IDセルがクリックできませんでした。")

def click_header_once(header_text="L-Gate連携ソースID"):
    xp = f"//table/thead//th[.//text()[contains(.,'{header_text}')]]"
    th = wait.until(EC.element_to_be_clickable((By.XPATH, xp)))
    safe_click(th)
    time.sleep(0.5)

def ensure_nonempty_first_row(col_header="L-Gate連携ソースID"):
    """先頭セルが空ならヘッダを1回だけクリックして『値ありが上』にする"""
    col_idx = get_column_index_by_header_text(col_header) or 8
    try:
        first_val = get_cell_text(1, col_idx)
    except:
        first_val = ""
    if first_val == "":
        click_header_once(col_header)

def first_time_filter_and_sort():
    """最初だけフィルタ＆ソート"""
    driver.get(SCHOOL_LIST_URL)
    time.sleep(0.8)
    name_input = find_by_label_input("学校名")
    if name_input:
        fill_and_fire(name_input, TARGET_SCHOOL_NAME)
    search_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='mainContens']//button[contains(text(),'検索')]")))
    try:
        driver.execute_script("arguments[0].removeAttribute('disabled');", search_btn)
    except:
        pass
    safe_click(search_btn)
    time.sleep(0.8)
    click_header_once("L-Gate連携ソースID")
    ensure_nonempty_first_row("L-Gate連携ソースID")

def wait_detail_screen_ready():
    wait.until(EC.visibility_of_element_located((By.ID, "mainContens")))
    time.sleep(0.2)

def find_lgate_input():
    """ data-bind に renkeimotoSourceId を持つ input を最優先で取得 """
    xps = [
        "//*[@id='mainContens']//input[contains(@data-bind,'renkeimotoSourceId')]",
        "(.//label[contains(normalize-space(.),'L-Gate連携ソースID')])[1]/following::input[contains(@class,'form-control')][1]",
        "//*[@id='mainContens']//input[contains(@class,'form-control') and (contains(@data-bind,'Gate') or contains(@name,'Gate') or contains(@placeholder,'L-Gate'))]",
    ]
    for xp in xps:
        try:
            el = wait.until(EC.presence_of_element_located((By.XPATH, xp)))
            if el.is_displayed():
                return el
        except:
            continue
    return None

def read_detail_school_name_and_id():
    """詳細画面で学校名と学校IDを取得（ログ用）"""
    name, schid = "", ""
    for xp in [
        "//*[@id='mainContens']//div[contains(@class,'page-heading-title')][1]",
        "//*[@id='mainContens']//input[@type='text' and contains(@data-bind,'gakkoName')]",
    ]:
        try:
            el = wait.until(EC.presence_of_element_located((By.XPATH, xp)))
            txt = (el.text or el.get_attribute("value") or "").strip()
            if txt:
                name = txt
                break
        except:
            continue
    try:
        el = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//*[@id='mainContens']//label[normalize-space(.)='学校ID']/following::input[1]")
        ))
        schid = (el.get_attribute("value") or "").strip()
    except:
        pass
    return name, schid

def clear_input_field_robust(el):
    """Knockout/React系でも確実に空にするための多段クリア（待機なし）"""
    try:
        driver.execute_script("arguments[0].removeAttribute('readonly');arguments[0].removeAttribute('disabled');", el)
    except:
        pass
    safe_click(el)
    try:
        el.clear()
    except:
        pass
    el.send_keys(Keys.CONTROL, 'a')
    el.send_keys(Keys.DELETE)
    driver.execute_script("arguments[0].value='';", el)
    for evt in ['input', 'keyup', 'change']:
        driver.execute_script("arguments[0].dispatchEvent(new Event(arguments[1], {bubbles:true}));", el, evt)
    driver.execute_script("document.activeElement && document.activeElement.blur();")
    try:
        WebDriverWait(driver, 5).until(lambda d: (el.get_attribute('value') or '').strip() == '')
    except:
        driver.execute_script("arguments[0].value='';", el)
        driver.execute_script("arguments[0].dispatchEvent(new Event('change',{bubbles:true}));", el)

def press_update_and_confirm_strict():
    """
    画面最下部の更新ボタンをクリック → ポップアップOK/はい（出たら）を押下。
    その後、**一覧のテーブルが見えるまで待機**（検索やり直しはしない）。
    """
    # 更新ボタン（実績XPath最優先）
    try:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(0.6)
        update_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//*[@id='mainContens']/div[12]/div[2]/button")
            )
        )
    except Exception:
        update_button = None
        for xp in [
            "//*[@id='mainContens']//button[contains(@data-bind,'updateGakkoInfo')]",
            "//*[@id='mainContens']//button[normalize-space(text())='更新']",
            "//button[contains(@class,'btn') and contains(.,'更新')]",
        ]:
            try:
                update_button = WebDriverWait(driver, 6).until(
                    EC.element_to_be_clickable((By.XPATH, xp))
                )
                break
            except:
                continue
        if not update_button:
            raise RuntimeError("更新ボタンが取得できませんでした。")

    driver.execute_script("arguments[0].scrollIntoView(true);", update_button)
    time.sleep(0.4)
    driver.execute_script("arguments[0].click();", update_button)

    # 確認ダイアログのOK/はい（出ない場合はスキップ）
    for xp in [
        "//button[normalize-space(text())='OK']",
        "//button[normalize-space(text())='Ok']",
        "//button[normalize-space(text())='はい']",
        "//*[@id='systemCommonConfirmDialogOk']",
    ]:
        try:
            popup_ok_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, xp))
            )
            driver.execute_script("arguments[0].click();", popup_ok_button)
            time.sleep(0.6)
            break
        except:
            continue

    # BlockUIローダ消失（出ない場合もある）
    try:
        WebDriverWait(driver, 20).until(
            EC.invisibility_of_element_located((By.ID, "ballSpinFadeLoaderBlockUI"))
        )
    except:
        pass

    # ★一覧に戻るまで待機（OK後は自動遷移する仕様）
    WebDriverWait(driver, 15).until(
        EC.visibility_of_element_located((By.XPATH, get_rows_xpath_base()))
    )

# ================= メイン処理 =================
try:
    # --- ログイン ---
    driver.get(BASE_URL)
    time.sleep(1.0)
    user_box = None
    for xp in [
        "/html/body/div/div[2]/form/div[1]/input",
        "//input[@type='text' or @name='userId' or contains(@placeholder,'ID')]",
    ]:
        try:
            user_box = wait.until(EC.presence_of_element_located((By.XPATH, xp)))
            break
        except:
            continue
    if not user_box:
        raise RuntimeError("ログインID入力欄が見つかりません。")

    pass_box = None
    for xp in [
        "/html/body/div/div[2]/form/div[2]/input",
        "//input[@type='password']",
    ]:
        try:
            pass_box = wait.until(EC.presence_of_element_located((By.XPATH, xp)))
            break
        except:
            continue
    if not pass_box:
        raise RuntimeError("パスワード入力欄が見つかりません。")

    fill_and_fire(user_box, LOGIN_ID)
    fill_and_fire(pass_box, LOGIN_PW)

    login_btn = None
    for xp in [
        "/html/body/div/div[2]/form/div[3]/button",
        "//button[contains(.,'ログイン') or contains(.,'Login') or contains(.,'ログオン')]",
        "//form//button",
    ]:
        try:
            login_btn = wait.until(EC.element_to_be_clickable((By.XPATH, xp)))
            break
        except:
            continue
    if not login_btn:
        raise RuntimeError("ログインボタンが見つかりません。")
    safe_click(login_btn)
    time.sleep(1.0)

    # --- 一覧 初回のみフィルタ＆ソート ---
    first_time_filter_and_sort()
    col_idx = get_column_index_by_header_text("L-Gate連携ソースID") or 8

    # === ループ：常に「先頭行」を処理 → OK後は自動で一覧に戻る ===
    while True:
        rows = get_rows()
        if not rows:
            print("行が見つからないため終了。")
            break

        try:
            first_val = get_cell_text(1, col_idx)
        except:
            first_val = ""

        if first_val == "":
            # 先頭が空なら値ありが枯渇 → 一度だけトグル、それでも空なら終了
            ensure_nonempty_first_row("L-Gate連携ソースID")
            try:
                first_val = get_cell_text(1, col_idx)
            except:
                first_val = ""
            if first_val == "":
                print("（一覧）先頭行の L-Gate連携ソースID が空です。処理対象がなくなったため終了。")
                break

        # 先頭の学校へ入る
        click_row_school_id(1)

        # 詳細画面：学校名/学校IDをログ用に取得 → L-Gate を空に → 更新（OK押下）→ 一覧へ自動復帰
        wait_detail_screen_ready()
        school_name, school_id = read_detail_school_name_and_id()

        input_field = find_lgate_input()
        if not input_field:
            raise RuntimeError("詳細画面の『L-Gate連携ソースID』入力欄が見つかりません。")

        clear_input_field_robust(input_field)
        press_update_and_confirm_strict()

        # ログ出力（どこが終わったか分かるように）
        print(f"完了: 学校名='{school_name}' / 学校ID='{school_id}' の L-Gate連携ソースIDを空にして保存しました。")

        # 一覧に戻っているので、そのまま次ループへ
        # 念のため、先頭が空ならヘッダを1回トグル（次の値ありを先頭へ）
        ensure_nonempty_first_row("L-Gate連携ソースID")

    print("=== 全処理完了 ===")

except Exception as e:
    print("エラー発生:", e)
    traceback.print_exc()

finally:
    cleanup_excel()
    # driver.quit()  # 自動で閉じたい場合はコメント解除
