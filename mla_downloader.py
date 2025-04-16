from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
import pandas as pd
import time
import glob
import os
from datetime import datetime

# --- Setup ---
print("🚀 Initialising...")
download_dir = os.path.abspath("downloads")
os.makedirs(download_dir, exist_ok=True)
cutoff_date = datetime.strptime("01/01/2024", "%d/%m/%Y")
log_file = "download_log.csv"

if os.path.exists(log_file):
    download_log = pd.read_csv(log_file)
    print(f"📓 Loaded existing log with {len(download_log)} entries.")
else:
    download_log = pd.DataFrame(columns=["Saleyard", "Report Date"])
    print("📓 Created new download log.")

options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "directory_upgrade": True
}
options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(options=options)

# --- Access site ---
print("🌐 Navigating to MLA PowerBI report...")
driver.get("https://next-app.nlrsreports.mla.com.au/saleyard-reports/cattle-prime")
time.sleep(20)

# --- Switch to iframe ---
iframes = driver.find_elements(By.TAG_NAME, "iframe")
for iframe in iframes:
    driver.switch_to.frame(iframe)
    if "powerbi" in driver.page_source.lower():
        break
print("🖼️ Switched into PowerBI iframe.")
time.sleep(3)

# --- Open Saleyard Dropdown ---
dropdown = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, "div[aria-label='Saleyard Name']"))
)
dropdown.click()
print("📂 Opened 'Saleyard Name' dropdown.")
time.sleep(2)

# --- Position mouse for scrolling ---
print("🖱️ Calculating dropdown location for scrolling...")
window_x = driver.execute_script("return window.screenX;")
window_y = driver.execute_script("return window.screenY;")
inner_x = driver.execute_script("return window.outerWidth - window.innerWidth;") // 2
header_y = driver.execute_script("return window.outerHeight - window.innerHeight;") - inner_x
location = dropdown.location
size = dropdown.size
element_x = window_x + location['x'] + size['width'] // 2 + inner_x
element_y = window_y + location['y'] + size['height'] + header_y + 90
pyautogui.moveTo(element_x, element_y)
time.sleep(1)

# --- Mouse helpers ---
def focus_on_saleyard_dropdown():
    pyautogui.moveTo(element_x, element_y)
    time.sleep(0.5)


def focus_on_report_date_dropdown():
    report_dropdown_x = element_x + 115
    pyautogui.moveTo(report_dropdown_x, element_y)
    time.sleep(0.5)
    




def relaunch_browser():
    global driver
    print("♻️ Relaunching browser...")
    driver.quit()
    os.system("pkill -f chromedriver")  # clean up any lingering processes
    time.sleep(2)
    driver = webdriver.Chrome(options=options)
    driver.get("https://next-app.nlrsreports.mla.com.au/saleyard-reports/cattle-prime")
    time.sleep(20)
    for iframe in driver.find_elements(By.TAG_NAME, "iframe"):
        driver.switch_to.frame(iframe)
        if "powerbi" in driver.page_source.lower():
            break
    print("🖼️ Relaunched and switched to PowerBI iframe.")
    focus_on_saleyard_dropdown()




# --- Scroll and collect saleyards ---
saleyard_names = []
print("🔽 Scrolling to collect saleyard names...")
for i in range(100):
    pyautogui.scroll(-1)
    time.sleep(0.6)
    items = driver.find_elements(By.CSS_SELECTOR, "div.slicerItemContainer span")
    new_names = 0
    for item in items:
        name = item.text.strip()
        if name and name not in saleyard_names:
            saleyard_names.append(name)
            print(f"  ➕ Found saleyard: {name}")
            new_names += 1
    if new_names == 0:
        print(f"  💤 No new names on scroll {i + 1}, stopping early.")
        break


print(f"🗂️ Total saleyards collected: {len(saleyard_names)}")

# --- Scroll back to the top ---
print("🔼 Scrolling back to top of saleyard list...")
for _ in range(20):
    pyautogui.scroll(+1)
    time.sleep(0.3)

# --- Reset dropdown after scrolling ---
print("🔁 Re-clicking dropdown to reset view...")
dropdown = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, "div[aria-label='Saleyard Name']"))
)
dropdown.click()
time.sleep(1)

previous_last_date = ""

# --- Process each saleyard in A–Z order ---
for yard_index, yard_name in enumerate(saleyard_names, 1):
    max_yard_retries = 2
    yard_retry = 0
    while yard_retry <= max_yard_retries:
        try:
            print(f"\n📍 ({yard_index}/{len(saleyard_names)}) Processing: {yard_name}")

            
            first_date = None
            skip_yard = False
            focus_on_saleyard_dropdown()
            time.sleep(0.5)
            # ✅ Re-fetch dropdown fresh inside loop
            dropdown = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "div[aria-label='Saleyard Name']"))
            )
            location = dropdown.location


            
            # --- Position mouse for scrolling ---
            print("🖱️ Calculating dropdown location for scrolling...")
            window_x = driver.execute_script("return window.screenX;")
            print("xx1xx")
            window_y = driver.execute_script("return window.screenY;")
            print("x2xxx")
            inner_x = driver.execute_script("return window.outerWidth - window.innerWidth;") // 2
            print("xxx3x")
            header_y = driver.execute_script("return window.outerHeight - window.innerHeight;") - inner_x
            print("xx4xx")
            location = dropdown.location
            print("xx5xx")
            size = dropdown.size
            print("xx6xx")
            element_x = window_x + location['x'] + size['width'] // 2 + inner_x
            print("xx7xx")
            element_y = window_y + location['y'] + size['height'] + header_y + 145
            print("xx8xx")
            pyautogui.moveTo(element_x, element_y)
            print("xx9xx")
            time.sleep(1)
            
            print("x20xxx")

            # Scroll downward to find saleyard
            found_yard = False
            print(f"🔎 Locating saleyard: {yard_name}")
            dropdown = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "div[aria-label='Saleyard Name']"))
            )
            dropdown.click()
            for scroll in range(20):
                items = driver.find_elements(By.CSS_SELECTOR, "div.slicerItemContainer span")
                for item in items:
                    current = item.text.strip()
                    name = item.text.strip()
                    print(f"  ➕ Found saleyard: {name}")
                    if current == yard_name:
                        item.click()
                        found_yard = True
                        print(f"✅ Clicked saleyard: {yard_name}")
                        break
                if found_yard:
                    break
                pyautogui.scroll(-2)
                time.sleep(0.6)

            if not found_yard:
                print(f"❌ Could not find saleyard: {yard_name}")
                continue

            time.sleep(2)


            report_dropdown = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "div[aria-label='Report Date']"))
            )
            report_dropdown.click()
            time.sleep(1)

            focus_on_report_date_dropdown()
            time.sleep(0.3)

            seen_dates = set()
            cutoff_reached = False

            print("🔽 Scrolling and processing report dates...")
            for i in range(100):
                if cutoff_reached or skip_yard:
                    break

                time.sleep(0.3)
                items = driver.find_elements(By.CSS_SELECTOR, "div.slicerItemContainer span")
                if not items:
                    print("❌ No report dates found in dropdown.")
                    break

                for item in items:
                    try:
                        raw_date = item.text.strip()
                    except Exception as e:
                        if "stale element reference" in str(e).lower():
                            print("♻️ Stale element encountered inside item loop. Triggering outer retry...")
                            raise  # ✅ Let the outer `except` handle the relaunch + retry
                        else:
                            raise e


                    if not raw_date or raw_date in seen_dates:
                        continue

                    if not first_date:
                        first_date = raw_date
                        if first_date == "(Blank)":
                            print("⚠️ First report date is '(Blank)'. Skipping saleyard.")
                            skip_yard = True
                            break
                        if first_date == previous_last_date:
                            print(f"⚠️ First report date {first_date} matches previous saleyard's last date. Skipping.")
                            skip_yard = True
                            break

                    seen_dates.add(raw_date)

                    try:
                        parsed_date = datetime.strptime(raw_date, "%d/%m/%Y")
                    except:
                        print(f"⚠️ Unreadable date format: {raw_date}")
                        continue

                    if parsed_date <= cutoff_date:
                        print(f"🛑 Reached old report date: {raw_date}, stopping saleyard.")
                        cutoff_reached = True
                        break

                    print(f"📅 Processing report date: {raw_date}")
                    driver.execute_script("arguments[0].scrollIntoView(true);", item)
                    time.sleep(0.5)
                    item.click()
                    time.sleep(2)

                    already_downloaded = not download_log[
                        (download_log["Saleyard"] == yard_name) &
                        (download_log["Report Date"] == raw_date)
                    ].empty

                    if already_downloaded:
                        print(f"📁 Already downloaded: {raw_date}")
                        continue

                    max_retries = 3
                    download_complete = False

                    for attempt in range(max_retries):
                        print(f"🔁 Attempt {attempt + 1} to export...")

                        try:
                            driver.switch_to.default_content()
                            export_btn = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Export Data')]"))
                            )
                            export_btn.click()
                            print("⬇️ Triggered export...")

                            download_start_time = time.time()
                            for _ in range(30):
                                time.sleep(1)
                                if glob.glob(os.path.join(download_dir, "*.crdownload")):
                                    continue

                                files = glob.glob(os.path.join(download_dir, "*.csv")) + glob.glob(os.path.join(download_dir, "*.xlsx"))
                                files = [f for f in files if os.path.getmtime(f) > download_start_time - 1]

                                if files:
                                    latest = max(files, key=os.path.getmtime)
                                    ext = os.path.splitext(latest)[1]
                                    new_name = f"{yard_name.replace(' ', '_')}_{raw_date.replace('/', '-')}{ext}"
                                    os.rename(latest, os.path.join(download_dir, new_name))
                                    print(f"✅ Saved: {new_name}")
                                    previous_last_date = raw_date  # ✅ only after successful save
                                    download_complete = True
                                    break

                            if download_complete:
                                break

                            print("⚠️ Download timed out, retrying...")

                        except Exception as e:
                            print(f"❌ Export attempt {attempt + 1} failed: {e}")

                    if not download_complete:
                        print(f"❌ Skipping {raw_date} after {max_retries} failed attempts.")
                        continue

                    download_log.loc[len(download_log)] = [yard_name, raw_date]
                    download_log.drop_duplicates(inplace=True)
                    download_log.to_csv(log_file, index=False)
                    print(f"📝 Log updated: {yard_name} - {raw_date}")

                    driver.switch_to.frame(driver.find_elements(By.TAG_NAME, "iframe")[0])
                    time.sleep(2)

                    driver.find_element(By.TAG_NAME, "body").click()
                    time.sleep(1)

                    report_dropdown = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "div[aria-label='Report Date']"))
                    )
                    report_dropdown.click()
                    time.sleep(1)

            if skip_yard:
                print(f"⛔ Skipping {yard_name} due to valid reason (blank or duplicate). No retry needed.")
                break  # ✅ Do not retry saleyard unless it's a stale element


            focus_on_saleyard_dropdown()
            
            break
            

        except Exception as e:
            if "stale element reference" in str(e).lower():
                print("♻️ Stale element encountered. Resetting dropdown and triggering retry...")
                yard_retry += 1
                relaunch_browser()
                continue  # Retry same saleyard
            

            elif "click intercepted" in str(e).lower() and yard_retry < max_yard_retries:
                print(f"⚠️ Click intercepted. Restarting browser and retrying {yard_name} (attempt {yard_retry + 1})...")
                yard_retry += 1
                relaunch_browser()
                continue  # Retry same saleyard

            else:
                print(f"⛔ Giving up on {yard_name} after {yard_retry + 1} attempt(s). Rebooting browser before continuing.")
                relaunch_browser()  # Ensure fresh state even after final failure
                break  # Move on to the next saleyard



driver.quit()
print("\n✅ All downloads complete.")