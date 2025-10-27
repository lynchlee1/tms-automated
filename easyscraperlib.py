import json
import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

import json
import os
import sys

class Settings:
    def __init__(self):
        self._system_data = self.load_system_constants()
    
    def _get_resource_path(self, relative_path):
        try: base_path = sys._MEIPASS
        except AttributeError: 
            current_dir = os.path.dirname(os.path.abspath(__file__))
            base_path = current_dir
        return os.path.join(base_path, relative_path)
    
    def load_system_constants(self):
        try:
            file_path = self._get_resource_path("system_constants.json")
            with open(file_path, 'r', encoding='utf-8') as f: return json.load(f)
        except Exception as e:
            print(f"Failed to load system_constants.json: {e}")
            return {}
    
    def get(self, key, default=None):
        for section in self._system_data.values():
            if isinstance(section, dict) and key in section: return section[key]
        return default
    
    def update(self, key, value):
        for section_name, section_data in self._system_data.items():
            if key in section_data:
                self._system_data[section_name][key] = value
                return True
        return False
    
    def get_section(self, section_name):
        return self._system_data.get(section_name, {})
    
    def update_section(self, section_name, data):
        try:
            self._system_data[section_name] = data
            return True
        except Exception: return False

_settings = Settings()
def get(key, default= None): return _settings.get(key, default)
def update(key, value): return _settings.update(key, value)
def get_section(section_name): return _settings.get_section(section_name)
def update_section(section_name, data): return _settings.update_section(section_name, data)
def get_resource_path(relative_path): return _settings._get_resource_path(relative_path)

class EasyScraper:
    def __init__(self, config, headless = False, process_type = None):
        self.config = config # from_date, to_date, company, key
        self.headless = headless
        self.driver = None
        self.wait = None

    def setup(self): 
        self.driver, self.wait = self._setup_driver(headless=self.headless)

    def cleanup(self): 
        if self.driver: self.driver.quit()
    
    def click_button(self, selector, in_iframe=False):
        try:
            # Switch to iframe popup if needed
            if in_iframe: self.driver.switch_to.frame(self.driver.find_element(By.CSS_SELECTOR, get("popup_iframe")))
            
            button = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
            self.driver.execute_script("arguments[0].click();", button)
            print(f"✅ CSS {selector} 버튼 클릭 완료")
            time.sleep(get("buffer_time"))
            
            if in_iframe: self.driver.switch_to.default_content()
                
        except Exception as e:
            if in_iframe: # Make sure to switch back to main frame on error
                try: self.driver.switch_to.default_content()
                except: pass
            raise Exception(f"❌ CSS {selector} 클릭 실패: {e}")

    def click_button_by_text(self, button_text, exact_match=True, in_iframe=False, element_type="auto"):
        """
        Find and click an element by its text content (searches all nested elements).
        Can search for buttons, links, or any clickable element.
        
        Args:
            button_text: The text to search for on the element
            exact_match: If True, matches exact text. If False, matches partial text.
            in_iframe: Whether the element is inside an iframe
        """
        try:
            if in_iframe:
                iframe_selector = get("popup_iframe")
                iframe = self.driver.find_element(By.CSS_SELECTOR, iframe_selector)
                self.driver.switch_to.frame(iframe)
            
            xpath_patterns = [
                f"//button[normalize-space(string())='{button_text}']",
                f"//a[normalize-space(string())='{button_text}']",
            ]
            for partial_xpath in xpath_patterns:
                try:
                    element = self.driver.find_element(By.XPATH, partial_xpath)
                    self.driver.execute_script("arguments[0].click();", element)
                    print(f"✅ {button_text} 클릭 완료")
                    if in_iframe:
                        self.driver.switch_to.default_content()
                        print(f"🔄 Switched back to main frame")
                    time.sleep(get("buffer_time"))
                    return
                except: continue
            raise Exception(f"{button_text} 클릭 실패")
                
        except Exception as e: 
            if in_iframe:
                try: self.driver.switch_to.default_content()
                except: pass
            raise Exception(f"❌ 요소 텍스트 '{button_text}' 클릭 실패: {e}")

    def fill_input(self, selector, value, in_iframe=False):
        if value is None: return
        try:
            # Switch to iframe if needed
            if in_iframe: self.driver.switch_to.frame(self.driver.find_element(By.CSS_SELECTOR, get("popup_iframe")))
                
            input = self.driver.find_element(By.CSS_SELECTOR, selector)
            self.driver.execute_script("arguments[0].click();", input) # Use JavaScript click to bypass popup overlay
            time.sleep(get("buffer_time"))

            input.clear()
            input.send_keys(value)
            time.sleep(get("buffer_time"))
            
            if in_iframe: self.driver.switch_to.default_content()
                
        except Exception as e: 
            if in_iframe: # Make sure to switch back to main frame on error
                try: self.driver.switch_to.default_content()
                except: pass
            raise Exception(f"❌ {selector} 입력 실패: {e}")
    
    def fill_dates(self):
        try:
            self._fill_input(get("from_date_selector"), self.config.get('from_date'))
            self._fill_input(get("to_date_selector"), self.config.get('to_date'))
        except Exception as e: raise Exception(f"❌ 날짜 입력 실패: {e}")

    def _setup_driver(self, headless):
        chrome_options = Options()
        if headless:
            chrome_options.add_argument("--headless=new")
        
        # Stability and crash prevention options
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--disable-web-security")
        chrome_options.add_argument("--disable-features=VizDisplayCompositor")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-plugins")
        chrome_options.add_argument("--disable-images")
        chrome_options.add_argument("--disable-default-apps")
        chrome_options.add_argument("--disable-sync")
        chrome_options.add_argument("--disable-translate")
        
        # Memory and performance options
        chrome_options.add_argument("--memory-pressure-off")
        chrome_options.add_argument("--max_old_space_size=4096")
        chrome_options.add_argument("--disable-background-networking")
        chrome_options.add_argument("--disable-background-timer-throttling")
        chrome_options.add_argument("--disable-backgrounding-occluded-windows")
        chrome_options.add_argument("--disable-renderer-backgrounding")
        
        # Logging and notifications
        chrome_options.add_argument("--disable-notifications")
        chrome_options.add_argument("--disable-logging")
        chrome_options.add_argument("--log-level=3")
        chrome_options.add_argument("--silent")
        chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        chrome_options.add_argument("--window-size=1600,1000")

        # Use Selenium Manager (built into Selenium 4.6+) for consistent driver resolution
        print("Chrome driver 시작 중... (Selenium Manager)")
        driver = webdriver.Chrome(options=chrome_options)
        driver.set_page_load_timeout(get("long_loadtime"))
        wait = WebDriverWait(driver, get("long_loadtime"))
        print("✅ Chrome driver 로딩 완료")
        return driver, wait
        