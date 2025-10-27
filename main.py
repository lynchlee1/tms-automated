from easyscraperlib import EasyScraper, get
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import pandas as pd
from datetime import datetime
import pyperclip

from config import USERID, PASSWORD

def scrape_once(config):
    print("Initializing scraper...")
    scraper = EasyScraper(config, headless=False, process_type='hist')
    scraper.setup()
    print("Opening details page...")
    scraper.driver.get(get("details_url"))
    time.sleep(get("buffer_time"))

    scraper.fill_input("#userId", USERID)
    scraper.fill_input("#password", PASSWORD)
    scraper.click_button("#root > div > div > div > div.login-right > div > form > button")
    time.sleep(get("buffer_time"))

    scraper.click_button_by_text("AI")
    time.sleep(get("buffer_time"))
    
    scraper.click_button_by_text("오퍼레이션")
    time.sleep(get("buffer_time"))

    weight_selector = "#root > div > div > div.flex-1.flex.flex-col.min-w-0.relative > main > div > div > div > div.container.mx-auto.p-4.space-y-4 > div:nth-child(5) > button:nth-child(2)"
    scraper.click_button(weight_selector)
    
    # Wait for table to appear after clicking weight_selector
    table_selector = "#root > div > div > div.flex-1.flex.flex-col.min-w-0.relative > main > div > div > div > div.container.mx-auto.p-4.space-y-4 > div.space-y-2 > div > span > div > div.datagrid.scroll > table"
    print("Waiting for table to appear...")
    table_elem = WebDriverWait(scraper.driver, get("long_loadtime")).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, table_selector))
    )
    print("Table appeared successfully!")
    thead = table_elem.find_element(By.TAG_NAME, "thead")
    header_rows = thead.find_elements(By.TAG_NAME, "tr")
    headers = []
    for row in header_rows:
        cells = row.find_elements(By.TAG_NAME, "th")
        row_data = [cell.text.strip() for cell in cells]
        if row_data:
            headers.append(row_data)
    
    # Flatten headers if nested
    if headers:
        # Combine multiple header rows if exists
        flat_headers = []
        max_cols = max(len(row) for row in headers) if headers else 0
        for col_idx in range(max_cols):
            header_parts = []
            for row in headers:
                if col_idx < len(row):
                    header_parts.append(row[col_idx])
            flat_headers.append(" - ".join(filter(None, header_parts)))
    
    # Extract table data from tbody
    tbody = table_elem.find_element(By.TAG_NAME, "tbody")
    rows = tbody.find_elements(By.TAG_NAME, "tr")
    
    data = []
    for row in rows:
        cells = row.find_elements(By.TAG_NAME, "td")
        row_data = [cell.text.strip() for cell in cells]
        if row_data:
            data.append(row_data)
    print(f"Extracted {len(data)} rows of data")
    df = pd.DataFrame(data, columns=flat_headers if headers else None)
    
    asset_selector = "#root > div > div > div.flex-1.flex.flex-col.min-w-0.relative > main > div > div > div > div.container.mx-auto.p-4.space-y-4 > div:nth-child(5) > button:nth-child(3)"
    scraper.click_button(asset_selector)
    
    # Wait for asset table to appear after clicking asset_selector
    print("Waiting for asset table to appear...")
    asset_table_elem = WebDriverWait(scraper.driver, get("short_loadtime")).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, table_selector))
    )
    print("Asset table appeared successfully!")
    
    # New approach: Click cell #cell0_d, right-click, select "Select All", Ctrl+C to copy all data
    print("Attempting to copy all table data via clipboard...")
    try:
        # Step 1: Click on cell #cell0_d
        cell0_d = WebDriverWait(scraper.driver, get("short_loadtime")).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#cell0_d"))
        )
        cell0_d.click()
        time.sleep(1)
        
        # Step 2: Right-click to open context menu
        action = ActionChains(scraper.driver)
        action.context_click(cell0_d).perform()
        time.sleep(1)
        
        # Step 3: Click "Select All" button in the hovering menu
        select_all_button = WebDriverWait(scraper.driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Select All')]"))
        )
        select_all_button.click()
        time.sleep(1)
        
        # Step 4: Copy to clipboard using Ctrl+C
        cell0_d.send_keys(Keys.chord(Keys.CONTROL, 'c'))
        time.sleep(2)
        
        # Step 5: Read clipboard data
        clipboard_data = pyperclip.paste()
        
        # Parse the tab-separated clipboard data into a DataFrame
        if clipboard_data:
            lines = clipboard_data.strip().split('\n')
            data_rows = [line.split('\t') for line in lines if line.strip()]
            if data_rows:
                # First row is usually headers
                headers2 = data_rows[0]
                data2 = data_rows[1:]  # Rest is data
                print(f"Extracted {len(data2)} rows of asset data from clipboard")
                df2 = pd.DataFrame(data2, columns=headers2)
            else:
                print("No data found in clipboard, falling back to regular scraping")
                raise Exception("No clipboard data")
        else:
            raise Exception("Clipboard is empty")
    except Exception as e:
        print(f"Clipboard method failed: {e}. Falling back to regular scraping...")
        # Fallback to original method
        thead2 = asset_table_elem.find_element(By.TAG_NAME, "thead")
        header_rows2 = thead2.find_elements(By.TAG_NAME, "tr")
        headers2 = []
        for row in header_rows2:
            cells = row.find_elements(By.TAG_NAME, "th")
            row_data = [cell.text.strip() for cell in cells]
            if row_data:
                headers2.append(row_data)
        
        # Flatten headers if nested
        if headers2:
            flat_headers2 = []
            max_cols = max(len(row) for row in headers2) if headers2 else 0
            for col_idx in range(max_cols):
                header_parts = []
                for row in headers2:
                    if col_idx < len(row):
                        header_parts.append(row[col_idx])
                flat_headers2.append(" - ".join(filter(None, header_parts)))
        
        # Extract asset table data from tbody
        tbody2 = asset_table_elem.find_element(By.TAG_NAME, "tbody")
        rows2 = tbody2.find_elements(By.TAG_NAME, "tr")
        
        data2 = []
        for row in rows2:
            cells = row.find_elements(By.TAG_NAME, "td")
            row_data = [cell.text.strip() for cell in cells]
            if row_data:
                data2.append(row_data)
        print(f"Extracted {len(data2)} rows of asset data")
        df2 = pd.DataFrame(data2, columns=flat_headers2 if headers2 else None)
    
    # Save both sheets to Excel
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_filename = f"results_{timestamp}.xlsx"
    
    with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='보유비중', index=False)
        df2.to_excel(writer, sheet_name='자산내역', index=False)
    
    print(f"Data saved to {excel_filename} in sheets '보유비중' and '자산내역'")
    

if __name__ == "__main__":
    scrape_once({})