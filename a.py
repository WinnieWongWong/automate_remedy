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

# Function to get the list of open tabs via Chrome's remote debugging protocol
def get_chrome_tabs(port=9222):
    try:
        response = requests.get(f'http://localhost:{port}/json')
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Error connecting to Chrome debugging port: {e}")
        return []

# Function to connect to an existing Chrome session
def connect_to_chrome_session(debugger_port=9222, target_url_prefix="https://localhost:43446/arsys/forms/argrp/SHR"):
    # Chrome options to attach to an existing session
    chrome_options = Options()

    chrome_options.add_experimental_option("debuggerAddress", f"localhost:{debugger_port}")

    # Initialize the WebDriver
    driver = webdriver.Chrome(options=chrome_options)

    # Get all open tabs
    tabs = get_chrome_tabs(debugger_port)
    target_window = None

    # Find the tab with the matching URL
    for tab in tabs:
        if 'url' in tab and tab['url'].startswith(target_url_prefix):
            target_window = tab['id']
            break

    if not target_window:
        print(f"No tab found with URL starting with {target_url_prefix}")
        driver.quit()
        return None

    # Switch to the target tab
    driver.switch_to.window(target_window)

    return driver

def goToView(debugger_port=9222, target_url_prefix="https://localhost:43446/arsys/forms/argrp/EGIS_CLP_SRMaster_ViewDetail"):
    
    chrome_options = Options()

    # Add debugger address option
    chrome_options.add_experimental_option("debuggerAddress", f"localhost:{debugger_port}")

    # Initialize the WebDriver
    driver = webdriver.Chrome(options=chrome_options)

    # Get all open tabs
    tabs = get_chrome_tabs(debugger_port)
    target_window = None

    # Find the tab with the matching URL
    for tab in tabs:
        if 'url' in tab and tab['url'].startswith(target_url_prefix):
            target_window = tab['id']
            break

    if not target_window:
        print(f"No tab found with URL starting with {target_url_prefix}")
        driver.quit()
        return None

    # Switch to the target tab
    driver.switch_to.window(target_window)

    return driver

# Main logic
def get_textarea_value(cr_value):

    base_path = Path("C:/remedy/2024/202409")
    # Connect to the Chrome session
    driver = connect_to_chrome_session()
    
    if driver is None:
        return
    
    try:
        if(cr_value != ""):

            changeIDtextarea = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "//label[text()='Change ID*+']/following-sibling::textarea"))
            )
            changeIDtextarea.clear()  # Clear existing text if needed
            changeIDtextarea.send_keys(cr_value) 

            searchBtn = driver.find_element(By.XPATH, "//a[div/div[text()='Search']]")
            driver.execute_script("arguments[0].click();", searchBtn)

            time.sleep(3)

            # Find the textarea element by its ID
            projectID = driver.find_element(By.XPATH, "//label[text()='Project ID']/following-sibling::textarea")

            # Get the value of the textarea
            projectIDValue = projectID.get_attribute("value")

            rText = driver.find_element(By.XPATH, "//label[text()='Summary*']/following-sibling::textarea")
            projectRText=rText.get_attribute("value")

            rrText = driver.find_element(By.XPATH, "//label[text()='Notes']/following-sibling::textarea")
            projectRRText=rrText.get_attribute("value")

            
            envText = driver.find_element(By.XPATH, "//label[text()='Site*']/following-sibling::textarea")
            projectEnvText= envText.get_attribute("value")

            ppTextBtn = driver.find_element(By.XPATH, "//a[.//div[text()='View']]")
            ppTextBtn.click()

            time.sleep(5)

            view_driver = goToView()

            if view_driver is None:
                return
            
            kp_excel_path = Path("C:/remedy/default/excel_file/CRQ000020241107_PROJ_PRD.xlsx")
            tt_excel_path = Path("C:/remedy/default/excel_file/CRQ000020241107_PROJ_TT.xlsx")

            #Alternative R content
            all_td_text = ""
            if not projectRRText.strip():
                rrText_td_elements = view_driver.find_elements(By.XPATH, "//div[@id='WIN_0_536871031']//table//tr//td")
                for td in rrText_td_elements:
                    all_td_text += td.text + "\n"  # Use "\n" for new lines instead of "/n"
                    print(td.text)
                projectRRText = all_td_text

            # Create the filename
            file_name_combine = f"{cr_value} {projectIDValue} {projectEnvText} {projectRText.strip().replace(":" , "").replace("|" , "")}"

            # Concatenate the path
            new_path = base_path / file_name_combine
        
            r_file = f"{new_path}/r.txt"

            try:
                new_path.mkdir(parents=True, exist_ok=True) #Create directory if it does not exist
                with open(r_file, 'w') as file:
                    file.write(projectRRText)
                if(projectEnvText == "TT"):
                    # Copy the file
                    tt_new_filename = f"{cr_value}_{projectIDValue}_TT.xlsx"
                    tt_new_full_path = new_path / tt_new_filename
                    shutil.copy(tt_excel_path, tt_new_full_path)
                    print(f'TT File copied to {tt_new_full_path}')
                else:
                    prod_new_filename = f"{cr_value}_{projectIDValue}_PRD.xlsx"
                    prod_new_full_path = new_path / prod_new_filename
                    shutil.copy(kp_excel_path, prod_new_full_path)
                    print(f'PRD File copied to {prod_new_full_path}')
            except Exception as e:
                print(f"An error occurred: {e}")

            
            file_name_combine = f"{cr_value} {projectIDValue} {projectEnvText} {projectRText}"
            destination_path = f"C:/remedy/2024/202409/{file_name_combine.strip().replace(":" , "").replace("|" , "")}"
            destination_path_c = Path(destination_path)

            documents_path = Path("C:/Users/user/Downloads")  # Replace with the path to your documents folder

            # Get list of files BEFORE clicking Save
            files_before = set(documents_path.iterdir())

            rows = view_driver.find_elements(By.XPATH, "//div[contains(@class, 'ardbnAttachmentPool')]//div[contains(@class, 'BaseTableOuter')]//div[contains(@class, 'BaseTableInner')]//table[@class='BaseTable']//tr")
            if not rows[1].find_element(By.XPATH, ".//td[1]/nobr/span").text:
                rows = view_driver.find_elements(By.XPATH, "//div[contains(@class, 'ardbnECFAttPool')]//div[contains(@class, 'BaseTableOuter')]//div[contains(@class, 'BaseTableInner')]//table[@class='BaseTable']//tr")
            else:
                print("There are some files will be downloaded, pls wait.....")
            for row in rows[1:]:  # Start from the second row to skip headers
                try:
                    col_one_element = row.find_element(By.XPATH, ".//td[1]/nobr/span")
                    col_one = col_one_element.text if col_one_element else None
                    print(col_one_element.text)
    
                    if col_one: 
                        actions = ActionChains(view_driver)
                        actions.click(row).perform()
                        selectBtn = view_driver.find_element(By.XPATH, "//a[text()='Save']")
                        selectBtn.click()
                        time.sleep(3) 

                    else:
                        print("End of clicking all files.Pls wait for the downloaded files move to destination.... ")
                        break
                except NoSuchElementException as ll:
                    print(f"Caught an exception: {ll}")
                except Exception as ee:
                    print(f"Caught an exception: {ee}")

            # Poll for new file(s) appearing in download folder
            timeout = 30  # Max wait time in seconds
            start_time = time.time()

            new_file = None
            while time.time() - start_time < timeout:
                files_after = set(documents_path.iterdir())
                new_files_set = files_after - files_before

                # Convert to list and filter out temporary/incomplete files
                new_files = [
                    f for f in new_files_set
                    if not f.name.endswith(('.tmp', '.crdownload', '.part', '.download'))  # Added '.download' for Edge
                    and f.is_file()
                ]
                        
                time.sleep(1)

            if new_files:
                for new_file in new_files: 
                    shutil.move(new_file, destination_path_c)
                    print(f"Successfully copied to: {destination_path_c / new_file.name}")
            else:
                print("No new file downloaded or timeout reached.")
    except Exception as ex:
        print(f"Caught an exception: {ex}")
    finally:
        driver.quit()
        

    
# Execute the script
if __name__ == "__main__":
    # Check if the correct number of arguments is provided
    if len(sys.argv) != 2:
        print("Usage: ./python a.py {xxxxx}")
        sys.exit(1)

    # Get the input parameter
    input_param = sys.argv[1]
    print(f"Input CR number: {input_param}")
    get_textarea_value(input_param)