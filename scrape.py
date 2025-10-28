from easyscraperlib import EasyScraper, get
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import glob
import os

import json
import sys

# Get USERID, PASSWORD, and headless from JSON file if provided, otherwise from config
def get_credentials_from_json(json_file):
    """Read USERID, PASSWORD, and headless mode from JSON file"""
    with open(json_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    userid = data.get('id', '')
    password = data.get('pw', '')
    # Check for headless mode - can be text "TRUE"/"FALSE" or boolean
    headless_str = str(data.get('headless', 'FALSE')).upper()
    headless = headless_str == 'TRUE'
    
    print(f"Read credentials from: {os.path.basename(json_file)}")
    print(f"Headless mode: {headless}")
    return userid, password, headless

def get_exe_dir():
    """Get the directory where the .exe is located"""
    if getattr(sys, 'frozen', False):
        # Running as compiled .exe
        return os.path.dirname(sys.executable)
    else:
        # Running as script
        return os.path.dirname(os.path.abspath(__file__))

# Check if JSON file is provided as command line argument
if len(sys.argv) > 1:
    # Get credentials from JSON file
    USERID, PASSWORD, HEADLESS_FROM_JSON = get_credentials_from_json(sys.argv[1])
else:
    # Fall back to config file
    try:
        from config import USERID, PASSWORD
        HEADLESS_FROM_JSON = None
    except ImportError:
        USERID = ""
        PASSWORD = ""
        HEADLESS_FROM_JSON = None

def convert_numeric_columns(df):
    """
    Convert columns containing pure numeric strings to numeric types
    
    df: pandas DataFrame
    -> pandas DataFrame with numeric conversions
    """
    df = df.copy()
    
    for col in df.columns:
        try:
            # Try to convert the column to numeric, forcing any non-numeric to NaN
            converted = pd.to_numeric(df[col], errors='coerce')
            
            # Check if all non-null values were successfully converted
            if df[col].notna().any() and (converted.notna().sum() / df[col].notna().sum()) > 0.8:
                # If more than 80% of non-null values are numeric, convert the column
                df[col] = converted
        except Exception:
            # If conversion fails, keep the original column as-is
            continue
    
    return df

def create_dataframe_from_rows(data_rows, headers):
    """
    Create a pandas DataFrame from datas.
    
    data_rows: list of row data (list of lists)
    headers: list of header names
    -> pandas.DataFrame with aligned headers
    """
    if not data_rows: return pd.DataFrame()

    num_cols = len(data_rows[0])
    if len(headers) < num_cols: headers = headers + [f'Column_{i+1}' for i in range(num_cols - len(headers))]
    elif len(headers) > num_cols: headers = headers[:num_cols]
    
    df = pd.DataFrame(data_rows, columns=headers)
    
    # Convert numeric columns
    df = convert_numeric_columns(df)
    
    return df

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

def scrape_table_to_clipboard(scraper, cell_selector):
    """
    cell_selector: CSS selector for the starting cell
    -> list of lists: Data rows parsed from clipboard

    * Waits until cell is loaded
    """
    try:
        cell_element = WebDriverWait(scraper.driver, get("long_loadtime")).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, cell_selector))
        )
        scraper.driver.execute_script("arguments[0].click();", cell_element)
        time.sleep(get("buffer_time"))
        
        action = ActionChains(scraper.driver)
        action.context_click(cell_element).perform()
        time.sleep(get("buffer_time"))
        
        scraper.click_button_by_text("Select All")
        time.sleep(get("buffer_time"))

        action = ActionChains(scraper.driver)
        action.context_click(cell_element).perform()
        time.sleep(get("buffer_time"))
        
        scraper.click_button_by_text("Copy Selected Cells")
        time.sleep(get("buffer_time"))
        
        data_rows = EasyScraper.parse_clipboard_to_rows()
        return data_rows
        
    except Exception as e:
        raise Exception(f"Error scraping clipboard data from {cell_selector}: {e}")

def scrape_table_to_clipboard_with_fallback(scraper, base_cell_selector, start_num=160, num_range=40, suffix="_Id"):
    """
    Try scraping with base_cell_selector, if fails, try alternative cell numbers
    
    base_cell_selector: e.g., "#cell105_Id"
    start_num: starting cell number
    num_range: how many cells above and below to try
    suffix: suffix for the cell selector (e.g., "_Id")
    """
    # Extract the selector pattern
    if "#cell" in base_cell_selector and suffix in base_cell_selector:
        # Try the original selector first
        try:
            return scrape_table_to_clipboard(scraper, base_cell_selector)
        except Exception as e1:
            print(f"Failed with {base_cell_selector}, trying alternatives...")
            
            # Try nearby cell numbers
            for offset in range(-num_range, num_range + 1):
                if offset == 0:  # Skip the original
                    continue
                    
                try_cell_num = start_num + offset
                try_selector = f"#cell{try_cell_num}{suffix}"
                
                try:
                    print(f"Trying {try_selector}...")
                    return scrape_table_to_clipboard(scraper, try_selector)
                except Exception:
                    continue
            
            # If all attempts fail, raise the original error
            raise Exception(f"Failed to find working cell selector after trying {num_range*2 + 1} alternatives: {e1}")
    else:
        # If pattern doesn't match, just try the original
        return scrape_table_to_clipboard(scraper, base_cell_selector)

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
    try:
        data_rows = scrape_table_to_clipboard(scraper, "#cell1_d")
        predefined_headers = ["날짜", "펀드 - 펀드", "AI(전략) - NAV", "MEZZ(전략) - 좌수", "AI + MEZZ - 평가액", "간접투자(전체) - 펀드내비중", "비시장성자산 - 평가액", "비유동성자산 - 펀드내비중", "평가액", "펀드내비중", "평가액", "펀드내비중", "평가액", "펀드내비중", "평가액", "펀드내비중"]
        df_weight = create_dataframe_from_rows(data_rows, predefined_headers)
    except Exception as e: raise Exception(f"Error processing weight data: {e}")

    scraper.click_button_by_text("자산내역")
    time.sleep(get("buffer_time"))
    try:
        data_rows = scrape_table_to_clipboard(scraper, "#cell0_d")
        predefined_headers = ["날짜", "펀드", "전략", "종목코드", "종목명", "매매제한", "보유수량", "종가", "직간접", "자산구분", "투자형태", "상장시장", "시가평가여부", "기초자산코드", "기초자산명", "기초자산구분", "기초자산투자형태", "기초자산 상장시장", "기초자산 기업코드", "기초자산 기업명", "기초자산기업 상장시장", "섹터"]
        df_asset = create_dataframe_from_rows(data_rows, predefined_headers)
        
        # Add calculated column: 평가액 = 보유수량 * 종가
        if "보유수량" in df_asset.columns and "종가" in df_asset.columns:
            # Convert to numeric if needed and multiply
            df_asset["평가액"] = pd.to_numeric(df_asset["보유수량"], errors='coerce') * pd.to_numeric(df_asset["종가"], errors='coerce')
            print(f"Added calculated column '평가액' (보유수량 * 종가)")
    except Exception as e: raise Exception(f"Error processing asset data: {e}")

    scraper.click_button_by_text("투자 원장 조회")
    time.sleep(get("buffer_time"))
    try:
        # Use fallback function to try alternative cell selectors if #cell105_Id fails
        data_rows = scrape_table_to_clipboard_with_fallback(scraper, "#cell105_Id", start_num=105, num_range=10, suffix="_Id")
        predefined_headers = ["ID", "자산코드", "자산명", "투자형태", "기초자산명", "구/신", "보유형태", "최초투자원금", "현재원금액", "현재평가액", "평가수익률", "회수수익률", "투자단가", "현재주가", "괴리율", "담당자(운용)", "담당자(지원)", "Exit예상(M)", "Exit예상(급)", "Exit방안(급)", "Exit예상(평)", "Exit방안(평)", "투자일", "전환가능일", "PUT최초일", "PUT다음일", "PUT최종일", "CALL최초일", "CALL종료일", "보호예수종료일", "만기일", "YTM", "YTP", "YTC", "CALL가능비율", "투자번호"]
        df_deal = create_dataframe_from_rows(data_rows, predefined_headers)
        print(f"Extracted {len(df_deal)} rows for 투자 원장")
    except Exception as e:
        print(f"Error processing 투자 원장 data: {e}")
        df_deal = pd.DataFrame()  # Create empty dataframe if there's an error

    # Save data to temp.xlsx in the same directory as the exe
    exe_dir = get_exe_dir()
    excel_filename = os.path.join(exe_dir, "temp.xlsx")
    print(f"Saving data to {excel_filename}")
    
    # Check if file exists, if not create a new Excel file with all sheets
    if os.path.exists(excel_filename):
        # For existing files, just use append mode with replace option
        with pd.ExcelWriter(excel_filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_weight.to_excel(writer, sheet_name='보유비중', index=False)
            df_asset.to_excel(writer, sheet_name='자산내역', index=False)
            df_deal.to_excel(writer, sheet_name='투자원장', index=False)
    else:
        with pd.ExcelWriter(excel_filename, engine='openpyxl', mode='w') as writer:
            df_weight.to_excel(writer, sheet_name='보유비중', index=False)
            df_asset.to_excel(writer, sheet_name='자산내역', index=False)
            df_deal.to_excel(writer, sheet_name='투자원장', index=False)
    print(f"Data saved to {excel_filename}")

if __name__ == "__main__":
    # Use headless from JSON if provided, otherwise check command line arguments
    if HEADLESS_FROM_JSON is not None:
        headless = HEADLESS_FROM_JSON
    else:
        headless = "--headless" in sys.argv or "-h" in sys.argv
    
    logging = "--logging" in sys.argv or "-l" in sys.argv
    scrape_once(headless=headless)