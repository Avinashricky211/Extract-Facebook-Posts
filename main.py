import os
import time
import json
import re
import random
import pickle
import logging
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from dotenv import load_dotenv
from datetime import datetime


def get_current_datetime():
    """Return current date and time formatted as 'DD/MM/YY hh:mm AM/PM'."""
    return datetime.now().strftime("%d/%m/%y %I:%M %p")


def add_scraping_datetime(result):
    """Append scraping datetime and TASKSLNO to result dictionary."""
    result["createdate"] = get_current_datetime()
    return result


def save_results_to_excel(results, filename='Facebook_Posts.xlsx'):
    """Save output results in Excel with specified columns and order."""
    df = pd.DataFrame([
        {
            "LIKES": r.get("like_count", 0),
            "COMMENTS": r.get("comment_count", 0),
            "SHARES": r.get("share_count", 0),
            "URL": r.get("post_url", ""),
            "CREATEDATE": r.get("createdate", "")
        }
        for r in results
    ])
    df = df[["LIKES", "COMMENTS", "SHARES", "URL", "CREATEDATE"]]
    df.to_excel(filename, index=False, engine='openpyxl')
    print(f"✅ Results saved to {filename}")


class FacebookDataCollector:
    def __init__(self):
        self.logger = self._setup_logger()
        self.logger.info("Initializing FacebookDataCollector")


        self.options = webdriver.ChromeOptions()
        #self.options.add_argument("--headless")
        self.options.add_argument("--disable-notifications")
        self.options.add_argument("--disable-blink-features=AutomationControlled")
        self.options.add_experimental_option("excludeSwitches", ["enable-automation"])
        self.options.add_experimental_option('useAutomationExtension', False)
        self.options.add_argument("--start-maximized")
        self.options.add_argument("--high-dpi-support=1")
        self.options.add_argument("--force-device-scale-factor=0.5")

        self.cookies_file = "facebook_cookies.pkl"
        self.driver = None
        self.wait = None
        self.driver = self._setup_driver()
        self.logger.info("Browser initialized successfully")

    def _setup_logger(self):
        logger = logging.getLogger('FacebookScraper')
        logger.setLevel(logging.INFO)

        for handler in logger.handlers[:]: 
            logger.removeHandler(handler)

        handler = logging.StreamHandler()
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        logger.addHandler(handler)

        return logger

    def _setup_driver(self):
        driver = webdriver.Chrome(options=self.options)
        driver.maximize_window()
        self.wait = WebDriverWait(driver, 30)
        return driver

    def _human_like_delay(self, min_delay=2, max_delay=5):
        """Add human-like delays between actions"""
        delay = random.uniform(min_delay, max_delay)
        time.sleep(delay)

    def _type_like_human(self, element, text):
        """Type text character by character with human-like delays"""
        element.clear()
        for char in text:
            element.send_keys(char)
            time.sleep(random.uniform(0.1, 0.3))

    def load_cookies(self):
        """Load Facebook cookies if available and valid"""
        try:
            if not self.driver:
                self.driver = self._setup_driver()
                
            if os.path.exists(self.cookies_file):
                self.driver.get("https://www.facebook.com")
                self._human_like_delay(3, 5)
                
                with open(self.cookies_file, 'rb') as f:
                    cookies = pickle.load(f)
                
                for cookie in cookies:
                    try:
                        self.driver.add_cookie(cookie)
                    except Exception as e:
                        self.logger.warning(f"Failed to add cookie: {e}")
                
                self.driver.refresh()
                self._human_like_delay(5, 8)
                
                if self._is_logged_in():
                    self.logger.info("Successfully logged in using cookies")
                    return True
                else:
                    self.logger.warning("Cookies appear to be invalid or expired")
                    return False
            else:
                self.logger.info("Cookies file not found")
                return False
                
        except Exception as e:
            self.logger.error(f"Error loading cookies: {e}")
            return False

    def _save_cookies(self):
        """Save current session cookies for future use"""
        try:
            cookies = self.driver.get_cookies()
            with open(self.cookies_file, 'wb') as f:
                pickle.dump(cookies, f)
            self.logger.info("Cookies saved successfully")
        except Exception as e:
            self.logger.error(f"Error saving cookies: {e}")

    def login_with_credentials(self):
        """Login using credentials from .env file with improved human-like behavior"""
        try:
            if not self.driver:
                self.driver = self._setup_driver()
                
            load_dotenv()
            username = os.getenv('FACEBOOK_EMAIL')
            password = os.getenv('FACEBOOK_PASSWORD')
            
            if not username or not password:
                raise ValueError("Facebook credentials not found in .env file")
            
            self.logger.info("Attempting login with credentials")
            
            self.driver.get("https://www.facebook.com")
            self._human_like_delay(5, 8)
            
            try:
                email_field = WebDriverWait(self.driver, 15).until(
                    EC.element_to_be_clickable((By.ID, "email"))
                )
                self._human_like_delay(1, 3)
                self._type_like_human(email_field, username)
                
                password_field = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "pass"))
                )
                self._human_like_delay(1, 3)
                self._type_like_human(password_field, password)
                
                self._human_like_delay(2, 4)
                
                login_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.NAME, "login"))
                )
                login_button.click()
                
                self._human_like_delay(10, 15)
                
                if self._is_logged_in():
                    self.logger.info("Successfully logged in with credentials")
                    self._save_cookies()
                    return True
                else:
                    self.logger.error("Login failed - please check credentials")
                    return False
                    
            except TimeoutException:
                self.logger.error("Timeout during login - elements not found")
                return False
                
        except Exception as e:
            self.logger.error(f"Error during credential login: {e}")
            return False

    def authenticate(self):
        """Primary authentication method - cookies first, then credentials"""
        self.logger.info("Starting authentication process")
        
        if self.load_cookies():
            return True
        
        self.logger.info("Falling back to credential authentication")
        return self.login_with_credentials()

    def login(self):
        """Main login method that uses enhanced authentication"""
        try:
            if not self.authenticate():
                raise Exception("Authentication failed")
            print("✅ Successfully logged in")
        except Exception as e:
            print(f"❌ Login failed: {e}")
            raise

    def _scroll_and_wait(self):
        """Simple scroll to ensure content is in view"""
        try:
            self.driver.execute_script("window.scrollTo(0, 100);")
            self._human_like_delay(2, 3)
        except Exception as e:
            self.logger.warning(f"Error during scrolling: {e}")

    def _wait_for_content_load(self):
        """Simple wait for page and metrics to load"""
        try:
            self._human_like_delay(3, 5)
        except Exception as e:
            self.logger.warning(f"Wait failed: {e}")

    def _is_logged_in(self):
        """Check if user is logged in to Facebook"""
        try:
            self._human_like_delay(3, 5)
            current_url = self.driver.current_url
            if "login" in current_url.lower():
                return False
            
            login_indicators = [
                "[data-testid='blue_bar_profile_link']",
                "[aria-label='Your profile']",
                "[data-testid='nav_search']",
                "div[role='banner']",
                "[aria-label='Facebook']"
            ]
            
            for indicator in login_indicators:
                try:
                    element = WebDriverWait(self.driver, 5).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, indicator))
                    )
                    if element:
                        return True
                except:
                    continue
            
            return False
            
        except Exception as e:
            self.logger.error(f"Error checking login status: {e}")
            return False

    def _parse_count(self, text):
        """Parse numeric values from text."""
        if not text:
            return 0
        try:
            txt = str(text).strip().lower()
            txt = re.sub(r'[^0-9.km]', '', txt)
            if not txt:
                return 0
            if txt.endswith('k'):
                return int(float(txt[:-1]) * 1000)
            if txt.endswith('m'):
                return int(float(txt[:-1]) * 1000000)
            return int(float(txt))
        except Exception:
            return 0

    def extract_metrics_with_xpath(self):
        """Extracts engagement metrics using XPath locators."""
        try:
            time.sleep(2)  
            
            metrics = {
                'like_count': 0,
                'comment_count': 0,
                'share_count': 0
            }
            
            import re
            
            try:
                js_code = """
                return document.evaluate(
                    '/html/body/div[1]/div/div[1]/div/div[5]/div/div/div[2]/div/div/div/div/div/div/div/div[2]/div[2]/div/div/div/div/div/div/div/div/div/div/div/div/div[13]/div/div/div[4]/div/div/div[1]/div/div[1]/div/div[1]/div/span/div/span[1]/span/span',
                    document,
                    null,
                    XPathResult.FIRST_ORDERED_NODE_TYPE,
                    null
                ).singleNodeValue?.textContent.trim() || '0';
                """
                likes_text = self.driver.execute_script(js_code)
                metrics['like_count'] = self._parse_count(likes_text)
                print(f"👍 Likes found: {metrics['like_count']}")
            except Exception as e:
                print(f"⚠️ Could not extract likes: {str(e)}")

            try:
                comments_js_code = """
                const xpath = '/html/body/div[1]/div/div[1]/div/div[5]/div/div/div[2]/div/div/div/div/div/div/div/div[2]/div[2]/div/div/div/div/div/div/div/div/div/div/div/div/div[13]/div/div/div[4]/div/div/div[1]/div/div[1]/div/div[2]/div[2]/span/div/span/span';
                
                const result = document.evaluate(
                    xpath,
                    document,
                    null,
                    XPathResult.FIRST_ORDERED_NODE_TYPE,
                    null
                ).singleNodeValue;

                if (result) {
                    const text = result.textContent.trim();
                    // Extract just the number from text like "39 comments" or "150 comments"
                    const match = text.match(/\\d+/);
                    if (match) {
                        return match[0];  // Return just the number
                    }
                }
                
                // Fallback to finding by aria-label if XPath fails
                const commentElement = Array.from(document.querySelectorAll('[aria-label*="comment" i]')).find(el => {
                    const match = el.textContent.match(/\\d+/);
                    return match !== null;
                });

                if (commentElement) {
                    const match = commentElement.textContent.match(/\\d+/);
                    return match ? match[0] : '0';
                }

                return '0';
                """
                comments_text = self.driver.execute_script(comments_js_code)
                metrics['comment_count'] = self._parse_count(comments_text)
                print(f"💬 Comments found: {metrics['comment_count']}")
            except Exception as e:
                print(f"⚠️ Could not extract comments: {str(e)}")

            try:
                shares_js_code = """
                const result = document.evaluate(
                    '/html/body/div[1]/div/div[1]/div/div[5]/div/div/div[2]/div/div/div/div/div/div/div/div[2]/div[2]/div/div/div/div/div/div/div/div/div/div/div/div/div[13]/div/div/div[4]/div/div/div[1]/div/div[1]/div/div[2]/div[3]/span/div/span/span',
                    document,
                    null,
                    XPathResult.FIRST_ORDERED_NODE_TYPE,
                    null
                ).singleNodeValue;

                if (result) {
                    const text = result.textContent.trim();
                    // Extract just the number from text like "5 shares" or "65 shares"
                    const match = text.match(/\\d+/);
                    if (match) {
                        return match[0];  // Return just the number
                    }
                }
                return '0';
                """
                shares_text = self.driver.execute_script(shares_js_code)
                metrics['share_count'] = int(shares_text)
                print(f"🔁 Shares found: {metrics['share_count']}")
            except Exception as e:
                print(f"⚠️ Could not extract shares: {str(e)}")

            print(f"✅ Successfully extracted metrics: {metrics}")
            return metrics

        except Exception as e:
            print(f"❌ Error extracting metrics: {str(e)}")
            return {'like_count': 0, 'comment_count': 0, 'share_count': 0}

    def process_single_photo(self, photo_url):
        """Process one Facebook post URL and return metrics + timestamp + TASKSLNO."""
        print(f"\n🔄 Processing: {photo_url}")
        self.driver.get(photo_url)
        self._wait_for_content_load()  

        data = {'post_url': photo_url}
        metrics = self.extract_metrics_with_xpath()
        data.update(metrics)
        data = add_scraping_datetime(data) 

        print(f"👍 Likes: {data['like_count']}, 💬 Comments: {data['comment_count']}, 🔁 Shares: {data['share_count']}")
        return data

    def close(self):
        if self.driver:
            self.driver.quit()
            print("🔴 Browser closed.")


def load_urls_from_file(filename="facebook_posts.txt"):
    if os.path.exists(filename):
        with open(filename) as f:
            return [line.strip() for line in f if line.strip()]
    return ["https://www.facebook.com"]


def process_multiple_urls(urls):
    collector = FacebookDataCollector()
    results = []
    try:
        collector.login()
        for i, url in enumerate(urls):
            if i > 0:  
                print(f"⏳ Waiting 7 seconds before processing next URL...")
                time.sleep(7)
            results.append(collector.process_single_photo(url))
    finally:
        collector.close()
    return results


def main():
    urls = load_urls_from_file()
    results = process_multiple_urls(urls)
    save_results_to_excel(results)


if __name__ == "__main__":
    main()