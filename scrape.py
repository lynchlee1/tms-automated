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
import glob
import os
from openpyxl import load_workbook

from config import USERID, PASSWORD
import sys

def find_latest_deallog_file():
    """Find the latest 메자닌_DealLog_{version}.xlsx file"""
    pattern = "메자닌_DealLog_*.xlsx"
    
    files = glob.glob(pattern)
    if not files: return None
    
    versions = []
    for file in files:
        # title should be 메자닌_DealLog_{version number} form
        try:
            version_str = os.path.splitext(file)[0].split("_")[-1]
            if version_str.isdigit(): versions.append((int(version_str), file))
        except: continue    
    if not versions: return None
    
    versions.sort(reverse=True)
    latest_file = versions[0][1]
    print(f"Found latest deal log file: {latest_file}")
    return latest_file

def scrape_once(headless=False):
    print("Initializing scraper...")
    scraper = EasyScraper(headless=headless)
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

    scraper.click_button_by_text("보유비중(AI,Bond,재간접)")
    time.sleep(get("buffer_time"))

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
    df_weight = pd.DataFrame(data, columns=flat_headers if headers else None)
    
    scraper.click_button_by_text("자산내역")
    time.sleep(get("short_loadtime"))
    
    try:
        cell0_d = WebDriverWait(scraper.driver, get("long_loadtime")).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#cell0_d"))
        )
        scraper.driver.execute_script("arguments[0].click();", cell0_d)
        time.sleep(get("buffer_time"))
        
        action = ActionChains(scraper.driver)
        action.context_click(cell0_d).perform()
        time.sleep(get("buffer_time"))
        
        scraper.click_button_by_text("Select All")
        time.sleep(get("buffer_time"))

        action = ActionChains(scraper.driver)
        action.context_click(cell0_d).perform()
        time.sleep(get("buffer_time"))
        
        scraper.click_button_by_text("Copy Selected Cells")
        time.sleep(get("buffer_time"))
        
        clipboard_data = pyperclip.paste()
        if clipboard_data:
            lines = clipboard_data.strip().split('\n')
            print(f"Found {len(lines)} lines in clipboard data")
            
            data_rows = []
            for line in lines:
                if line.strip():
                    row = line.split('\t')
                    data_rows.append(row)                      
            num_cols = len(data_rows[0])
            headers = ["날짜", "펀드", "전략", "종목코드", "종목명", "매매제한", "보유수량", "종가", "직간접", "자산구분", "투자형태", "상장시장", "시가평가여부", "기초자산코드", "기초자산명", "기초자산구분", "기초자산투자형태", "기초자산 상장시장", "기초자산 기업코드", "기초자산 기업명", "기초자산기업 상장시장", "섹터"]
            if len(headers) < num_cols:
                headers = headers + [f'Column_{i+1}' for i in range(num_cols - len(headers))]
            headers = headers[:num_cols]

            df_asset = pd.DataFrame(data_rows, columns=headers)
        else: raise Exception("No data found in clipboard")
    except Exception as e: raise Exception(f"Error processing asset data: {e}")

    # Save data to temp.xlsx
    excel_filename = "temp.xlsx"
    print(f"Saving data to {excel_filename}")
    
    # Check if file exists, if not create a new Excel file with both sheets
    if os.path.exists(excel_filename):
        book = load_workbook(excel_filename)
        with pd.ExcelWriter(excel_filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            writer.book = book
            df_weight.to_excel(writer, sheet_name='보유비중', index=False)
            df_asset.to_excel(writer, sheet_name='자산내역', index=False)
    else:
        with pd.ExcelWriter(excel_filename, engine='openpyxl', mode='w') as writer:
            df_weight.to_excel(writer, sheet_name='보유비중', index=False)
            df_asset.to_excel(writer, sheet_name='자산내역', index=False)
    print(f"Data saved to {excel_filename}")

if __name__ == "__main__":
    headless = "--headless" in sys.argv or "-h" in sys.argv
    logging = "--logging" in sys.argv or "-l" in sys.argv
    
    scrape_once(headless=headless)