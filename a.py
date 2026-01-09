import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import requests
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from pathlib import Path
import shutil
import sys

# ==================== GLOBAL CONFIGURATION ====================
DEBUGGER_PORT = 9222

MAIN_URL_PREFIX = "https://localhost:43446/arsys/forms/argrp/SHR"
VIEW_URL_PREFIX = "https://localhost:43446/arsys/forms/argrp/EGIS_CLP_SRMaster_ViewDetail"

BASE_SAVE_PATH = Path("C:/remedy/2024/202409")
DOWNLOADS_PATH = Path("C:/Users/user/Downloads")

KP_EXCEL_TEMPLATE = Path("C:/remedy/default/excel_file/CRQ000020241107_PROJ_PRD.xlsx")
TT_EXCEL_TEMPLATE = Path("C:/remedy/default/excel_file/CRQ000020241107_PROJ_TT.xlsx")
# ============================================================

def get_chrome_tabs(port=DEBUGGER_PORT):
    try:
        response = requests.get(f'http://localhost:{port}/json')
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Error connecting to Chrome debugging port {port}: {e}")
        return []

def connect_to_chrome_session(target_url_prefix=MAIN_URL_PREFIX, port=DEBUGGER_PORT):
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", f"localhost:{port}")
    driver = webdriver.Chrome(options=chrome_options)

    tabs = get_chrome_tabs(port)
    target_window = None
    for tab in tabs:
        if 'url' in tab and tab['url'].startswith(target_url_prefix):
            target_window = tab['id']
            break

    if not target_window:
        print(f"No tab found with URL starting with {target_url_prefix}")
        driver.quit()
        return None

    driver.switch_to.window(target_window)
    return driver

def go_to_view(target_url_prefix=VIEW_URL_PREFIX, port=DEBUGGER_PORT):
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", f"localhost:{port}")
    driver = webdriver.Chrome(options=chrome_options)

    tabs = get_chrome_tabs(port)
    target_window = None
    for tab in tabs:
        if 'url' in tab and tab['url'].startswith(target_url_prefix):
            target_window = tab['id']
            break

    if not target_window:
        print(f"No tab found with URL starting with {target_url_prefix}")
        driver.quit()
        return None

    driver.switch_to.window(target_window)
    return driver

def get_textarea_value(cr_value):
    if not cr_value:
        print("CR value is empty.")
        return

    driver = connect_to_chrome_session()
    if driver is None:
        return

    try:
        # Input CR and search
        changeIDtextarea = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//label[text()='Change ID*+']/following-sibling::textarea"))
        )
        changeIDtextarea.clear()
        changeIDtextarea.send_keys(cr_value)

        searchBtn = driver.find_element(By.XPATH, "//a[div/div[text()='Search']]")
        driver.execute_script("arguments[0].click();", searchBtn)
        time.sleep(3)

        # Extract values
        projectIDValue = driver.find_element(By.XPATH, "//label[text()='Project ID']/following-sibling::textarea").get_attribute("value")
        projectRText = driver.find_element(By.XPATH, "//label[text()='Summary*']/following-sibling::textarea").get_attribute("value")
        projectRRText = driver.find_element(By.XPATH, "//label[text()='Notes']/following-sibling::textarea").get_attribute("value")
        projectEnvText = driver.find_element(By.XPATH, "//label[text()='Site*']/following-sibling::textarea").get_attribute("value")

        # Click View button
        view_btn = driver.find_element(By.XPATH, "//a[.//div[text()='View']]")
        view_btn.click()
        time.sleep(5)

        view_driver = go_to_view()
        if view_driver is None:
            return

        # Fallback: scrape Notes from table if empty
        if not projectRRText.strip():
            td_elements = view_driver.find_elements(By.XPATH, "//div[@id='WIN_0_536871031']//table//tr//td")
            projectRRText = "\n".join(td.text for td in td_elements if td.text.strip())

        # Create folder and save r.txt
        safe_title = projectRText.strip().replace(":", "").replace("|", "")
        folder_name = f"{cr_value} {projectIDValue} {projectEnvText} {safe_title}"
        new_path = BASE_SAVE_PATH / folder_name
        new_path.mkdir(parents=True, exist_ok=True)

        with open(new_path / "r.txt", "w", encoding="utf-8") as f:
            f.write(projectRRText)

        # Copy appropriate Excel template
        if projectEnvText == "TT":
            target_excel = TT_EXCEL_TEMPLATE
            suffix = "_TT.xlsx"
        else:
            target_excel = KP_EXCEL_TEMPLATE
            suffix = "_PRD.xlsx"

        excel_dest = new_path / f"{cr_value}_{projectIDValue}{suffix}"
        shutil.copy(target_excel, excel_dest)
        print(f"Excel copied to: {excel_dest}")

        # Download attachments
        files_before = set(DOWNLOADS_PATH.iterdir())

        rows = view_driver.find_elements(By.XPATH, "//div[contains(@class, 'ardbnAttachmentPool')]//table[@class='BaseTable']//tr")
        # Fallback to alternative attachment table
        if len(rows) < 2 or not rows[1].find_element(By.XPATH, ".//td[1]/nobr/span").text.strip():
            rows = view_driver.find_elements(By.XPATH, "//div[contains(@class, 'ardbnECFAttPool')]//table[@class='BaseTable']//tr")

        print("Starting download of attachments...")
        for row in rows[1:]:
            try:
                filename = row.find_element(By.XPATH, ".//td[1]/nobr/span").text.strip()
                if not filename:
                    break
                ActionChains(view_driver).click(row).perform()
                save_btn = view_driver.find_element(By.XPATH, "//a[text()='Save']")
                save_btn.click()
                time.sleep(3)
            except Exception as e:
                print(f"Error processing row: {e}")

        # Wait for downloads to complete
        timeout = 60
        start_time = time.time()
        downloaded_files = []
        while time.time() - start_time < timeout:
            files_after = set(DOWNLOADS_PATH.iterdir())
            new_files = files_after - files_before
            downloaded_files = [
                f for f in new_files
                if f.is_file() and not f.name.endswith(('.tmp', '.crdownload', '.part', '.download'))
            ]
            if downloaded_files:
                break
            time.sleep(2)

        # Move downloaded files to target folder
        if downloaded_files:
            for f in downloaded_files:
                shutil.move(str(f), str(new_path / f.name))
                print(f"Moved: {f.name} → {new_path}")
        else:
            print("No new files downloaded or timeout reached.")

    except Exception as ex:
        print(f"Unexpected error: {ex}")
    finally:
        driver.quit()
        if 'view_driver' in locals():
            view_driver.quit()

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python a.py <CR_NUMBER>")
        sys.exit(1)

    cr_value = sys.argv[1].strip()
    print(f"Processing CR: {cr_value}")
    get_textarea_value(cr_value)