"""
auto_apply.py
─────────────
Smart LinkedIn Easy Apply bot with multiple detection strategies.
"""

import argparse
import json
import os
import time
import datetime
import traceback
from pathlib import Path

import openpyxl

try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.chrome.service import Service
    from selenium.common.exceptions import TimeoutException, NoSuchElementException
    SELENIUM_AVAILABLE = True
except ImportError:
    SELENIUM_AVAILABLE = False
    print("Run: pip install selenium webdriver-manager")


def get_driver() -> "webdriver.Chrome":
    from selenium.webdriver.chrome.options import Options
    try:
        from webdriver_manager.chrome import ChromeDriverManager
        service = Service(ChromeDriverManager().install())
    except ImportError:
        service = Service()

    opts = Options()
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    opts.add_argument("--window-size=1440,900")
    opts.add_argument("--start-maximized")

    driver = webdriver.Chrome(service=service, options=opts)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver


def find_easy_apply_button(driver):
    """Try every possible way to find the Easy Apply button."""
    strategies = [
        (By.XPATH, '//button[contains(.,"Easy Apply")]'),
        (By.XPATH, '//button[contains(@aria-label,"Easy Apply")]'),
        (By.XPATH, '//button[contains(@class,"jobs-apply-button")]'),
        (By.XPATH, '//button[contains(.,"easy apply")]'),
        (By.XPATH, '//button[contains(.,"Easy apply")]'),
        (By.CSS_SELECTOR, '.jobs-apply-button'),
        (By.CSS_SELECTOR, '[data-control-name="jobdetails_topcard_inapply"]'),
        (By.XPATH, '//div[contains(@class,"jobs-apply")]//button'),
        (By.XPATH, '//button[@data-job-id]'),
    ]

    for by, selector in strategies:
        try:
            buttons = driver.find_elements(by, selector)
            for btn in buttons:
                if btn.is_displayed() and btn.is_enabled():
                    text = btn.text.lower()
                    aria = (btn.get_attribute("aria-label") or "").lower()
                    if "easy" in text or "easy" in aria or "apply" in text:
                        return btn
        except:
            continue
    return None


def click_easy_apply(driver):
    """Find and click Easy Apply button using JavaScript as fallback."""
    btn = find_easy_apply_button(driver)
    if btn:
        try:
            driver.execute_script("arguments[0].scrollIntoView(true);", btn)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", btn)
            print("✅ Clicked Easy Apply button!")
            return True
        except:
            pass

    # JavaScript search as last resort
    try:
        result = driver.execute_script("""
            var buttons = document.querySelectorAll('button');
            for(var i=0; i<buttons.length; i++){
                var txt = buttons[i].innerText.toLowerCase();
                var aria = (buttons[i].getAttribute('aria-label') || '').toLowerCase();
                if(txt.includes('easy apply') || aria.includes('easy apply')){
                    buttons[i].click();
                    return true;
                }
            }
            return false;
        """)
        if result:
            print("✅ Clicked Easy Apply via JavaScript!")
            return True
    except:
        pass

    return False


def wait_for_modal(driver, timeout=30):
    """Wait for Easy Apply modal to appear."""
    selectors = [
        '//div[contains(@class,"jobs-easy-apply-modal")]',
        '//div[contains(@class,"easy-apply-modal")]',
        '//h2[contains(.,"Apply to")]',
        '//h2[contains(.,"Easy Apply")]',
        '//button[contains(@aria-label,"Submit application")]',
        '//button[contains(@aria-label,"Review your application")]',
    ]
    for _ in range(timeout):
        for selector in selectors:
            try:
                el = driver.find_element(By.XPATH, selector)
                if el.is_displayed():
                    return True
            except:
                pass
        time.sleep(1)
    return False


def fill_form_step(driver, config, resume_path):
    """Fill all fields in current form step."""
    # Upload resume
    try:
        file_inputs = driver.find_elements(By.XPATH, '//input[@type="file"]')
        for fi in file_inputs:
            if fi.is_displayed() or True:
                fi.send_keys(os.path.abspath(resume_path))
                time.sleep(2)
                break
    except:
        pass

    # Fill phone number
    for selector in [
        '//input[contains(@id,"phoneNumber")]',
        '//input[contains(@name,"phoneNumber")]',
        '//input[@name="phone"]',
        '//input[contains(@placeholder,"Phone")]',
        '//input[contains(@placeholder,"phone")]',
    ]:
        try:
            el = driver.find_element(By.XPATH, selector)
            if el.is_displayed():
                el.clear()
                el.send_keys(config.get("phone", ""))
                break
        except:
            pass

    # Fill text fields by label
    label_value_map = {
        "phone":      config.get("phone", ""),
        "mobile":     config.get("phone", ""),
        "city":       config.get("location", ""),
        "location":   config.get("location", ""),
        "year":       config.get("years_experience", ""),
        "experience": config.get("years_experience", ""),
        "salary":     config.get("salary_expected", ""),
        "expected":   config.get("salary_expected", ""),
    }

    try:
        labels = driver.find_elements(By.XPATH, '//label')
        for label in labels:
            label_text = label.text.lower()
            for keyword, value in label_value_map.items():
                if keyword in label_text and value:
                    try:
                        input_id = label.get_attribute("for")
                        if input_id:
                            inp = driver.find_element(By.ID, input_id)
                            if inp.get_attribute("type") not in ("file", "hidden", "checkbox", "radio"):
                                inp.clear()
                                inp.send_keys(value)
                    except:
                        pass
    except:
        pass

    # Handle Yes/No radio buttons
    yes_no_map = {
        "authorized":  config.get("work_auth", "Yes"),
        "sponsorship": config.get("requires_sponsor", "No"),
        "visa":        config.get("requires_sponsor", "No"),
        "relocate":    "Yes",
        "remote":      "Yes",
    }
    try:
        fieldsets = driver.find_elements(By.TAG_NAME, "fieldset")
        for fieldset in fieldsets:
            try:
                legend = fieldset.find_element(By.TAG_NAME, "legend").text.lower()
                answer = "Yes"
                for keyword, ans in yes_no_map.items():
                    if keyword in legend:
                        answer = ans
                        break
                radio = fieldset.find_element(
                    By.XPATH,
                    f'.//label[contains(.,"{answer}")]/..//input[@type="radio"] | '
                    f'.//input[@type="radio"][following-sibling::*[contains(.,"{answer}")]] | '
                    f'.//input[@type="radio"][@value="{answer}"]'
                )
                driver.execute_script("arguments[0].click();", radio)
            except:
                pass
    except:
        pass


def advance_form(driver):
    """Click Next/Review/Continue button. Returns True if advanced, False if stuck."""
    for btn_text in ["Next", "Review", "Continue to next step", "Continue"]:
        try:
            btns = driver.find_elements(
                By.XPATH,
                f'//button[contains(@aria-label,"{btn_text}")] | //button[contains(.,"{btn_text}")]'
            )
            for btn in btns:
                if btn.is_displayed() and btn.is_enabled():
                    driver.execute_script("arguments[0].click();", btn)
                    time.sleep(2)
                    return True
        except:
            pass
    return False


def submit_application(driver):
    """Try to click Submit button."""
    for selector in [
        '//button[contains(@aria-label,"Submit application")]',
        '//button[contains(.,"Submit application")]',
        '//button[contains(.,"Submit")]',
    ]:
        try:
            btns = driver.find_elements(By.XPATH, selector)
            for btn in btns:
                if btn.is_displayed() and btn.is_enabled():
                    driver.execute_script("arguments[0].click();", btn)
                    time.sleep(3)
                    print("✅ Application submitted!")
                    return True
        except:
            pass
    return False


class LinkedInBot:
    def __init__(self, driver, config):
        self.driver = driver
        self.cfg = config

    def login(self):
        self.driver.get("https://www.linkedin.com/login")
        time.sleep(3)
        try:
            self.driver.find_element(By.ID, "username").send_keys(self.cfg["linkedin_email"])
            time.sleep(1)
            self.driver.find_element(By.ID, "password").send_keys(self.cfg["linkedin_password"])
            time.sleep(1)
            self.driver.find_element(By.XPATH, '//button[@type="submit"]').click()
            time.sleep(5)
            print("✅ LinkedIn: Logged in")
        except Exception as e:
            print(f"⚠️ Login issue: {e}")

    def apply(self, job_url, resume_path):
        self.driver.get(job_url)
        time.sleep(6)

        print("🔍 Looking for Easy Apply button...")

        # Try auto-click first
        clicked = click_easy_apply(self.driver)

        if not clicked:
            print("\n" + "="*55)
            print("👆 Could not auto-click. PLEASE CLICK 'Easy Apply' NOW")
            print("   You have 60 seconds...")
            print("="*55 + "\n")

        # Wait for modal
        print("⏳ Waiting for application form to open...")
        modal_found = wait_for_modal(self.driver, timeout=60)

        if not modal_found:
            print("❌ Form not detected. Try clicking Easy Apply manually.")
            return False

        print("✅ Form opened! Filling in your details...")

        # Walk through form
        for step in range(10):
            time.sleep(2)
            fill_form_step(self.driver, self.cfg, resume_path)

            if submit_application(self.driver):
                return True

            if not advance_form(self.driver):
                print(f"⚠️ On step {step+1} - please continue manually if needed")
                time.sleep(10)
                # Try submit one more time
                if submit_application(self.driver):
                    return True
                break

        return False


def update_tracker_status(tracker_path, job_url, status, notes=""):
    try:
        wb = openpyxl.load_workbook(tracker_path)
        ws = wb["Applications"]
        for row in ws.iter_rows(min_row=2, max_row=200):
            url_cell = row[10]
            if url_cell.value and job_url in str(url_cell.value):
                row[7].value = status
                row[18].value = (row[18].value or "") + f" | {notes}"
                wb.save(tracker_path)
                print(f"✅ Tracker updated: {status}")
                return
        print("⚠️ URL not found in tracker")
    except Exception as e:
        print(f"⚠️ Tracker error: {e}")


def main():
    if not SELENIUM_AVAILABLE:
        print("Install: pip install selenium webdriver-manager")
        return

    parser = argparse.ArgumentParser()
    parser.add_argument("--url",      required=True)
    parser.add_argument("--platform", required=True, choices=["linkedin", "generic"])
    parser.add_argument("--resume",   required=True)
    parser.add_argument("--tracker",  required=True)
    parser.add_argument("--config",   required=True)
    args = parser.parse_args()

    config = json.loads(Path(args.config).read_text())
    driver = get_driver()

    try:
        if args.platform == "linkedin":
            bot = LinkedInBot(driver, config)
            bot.login()
            success = bot.apply(args.url, args.resume)
        else:
            print("Use platform=generic for non-LinkedIn jobs")
            success = False

        status = "Applied" if success else "Manual Review Needed"
        update_tracker_status(args.tracker, args.url, status)

    except Exception as e:
        print(f"❌ Error: {e}")
        traceback.print_exc()
        update_tracker_status(args.tracker, args.url, "Error", str(e))

    finally:
        print("\n✅ Done! Closing browser in 5 seconds...")
        time.sleep(5)
        driver.quit()


if __name__ == "__main__":
    main()
