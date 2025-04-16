from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
import time
import glob
import os

download_dir = os.path.abspath("downloads")  # Create a folder in your current dir

options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "directory_upgrade": True
}
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=options)


driver.get("https://next-app.nlrsreports.mla.com.au/saleyard-reports/cattle-prime")
time.sleep(10)

# Switch into iframe
iframes = driver.find_elements(By.TAG_NAME, "iframe")
for iframe in iframes:
    driver.switch_to.frame(iframe)
    if "powerbi" in driver.page_source.lower():
        break

time.sleep(3)

# Click 'Saleyard Name' dropdown
try:
    dropdown = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "div[aria-label='Saleyard Name']"))
    )
    dropdown.click()
    print("✅ Clicked dropdown.")
    time.sleep(2)
except Exception as e:
    print(f"❌ Failed to click dropdown: {e}")
    driver.quit()
    exit()

# Get screen offset of browser window
window_x = driver.execute_script("return window.screenX;")
window_y = driver.execute_script("return window.screenY;")
inner_x = driver.execute_script("return window.outerWidth - window.innerWidth;") // 2
header_y = driver.execute_script("return window.outerHeight - window.innerHeight;") - inner_x

# Get element location
location = dropdown.location
size = dropdown.size
element_x = window_x + location['x'] + size['width'] // 2 + inner_x
element_y = window_y + location['y'] + size['height'] + header_y + 90

# Move mouse to dropdown area
pyautogui.moveTo(element_x, element_y)
time.sleep(1)

# Scroll to find 'Scone'
found_scone = False
seen_saleyards = set()

for i in range(100):
    print(f"🌀 Scroll attempt {i + 1}")
    pyautogui.scroll(-1)
    time.sleep(0.7)

    visible_items = driver.find_elements(By.CSS_SELECTOR, "div.slicerItemContainer span")
    for item in visible_items:
        name = item.text.strip()
        if name and name not in seen_saleyards:
            seen_saleyards.add(name)
            print(f"👁️  Found: {name}")
        if name == "Scone":
            item.click()
            print("✅ Clicked 'Scone'.")
            found_scone = True
            break

    if found_scone:
        break

if not found_scone:
    print("❌ Couldn't find 'Scone'")
    print(f"📝 Saleyards seen ({len(seen_saleyards)}): {sorted(seen_saleyards)}")
    driver.quit()
    exit()

# Wait for report dates to load
time.sleep(5)

# Click 'Report Date' dropdown
try:
    report_dropdown = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "div[aria-label='Report Date']"))
    )
    report_dropdown.click()
    print("📅 Clicked report date dropdown.")
    time.sleep(2)
except Exception as e:
    print(f"❌ Failed to click report date dropdown: {e}")
    driver.quit()
    exit()

# Log and select first & second available report dates
try:
    date_items = [item for item in driver.find_elements(By.CSS_SELECTOR, "div.slicerItemContainer span") if item.text.strip()]
    for i, item in enumerate(date_items):
        text = item.text.strip()
        print(f"🗓️  Found report date: {text}")
    
    if len(date_items) >= 2:
        date_items[1].click()
        print(f"🔁 Temporarily selected second date: {date_items[1].text.strip()}")
        time.sleep(1)
        date_items[0].click()
        print(f"✅ Re-selected most recent date: {date_items[0].text.strip()}")
        found_report = True
    elif len(date_items) == 1:
        date_items[0].click()
        print(f"✅ Only one date found. Selected: {date_items[0].text.strip()}")
        found_report = True
    else:
        print("❌ No report dates found.")
        found_report = False
except Exception as e:
    print(f"❌ Failed during date selection logic: {e}")
    found_report = False


# Click 'Export data' button
# Click 'Export Data' button
if found_report:
    time.sleep(3)
    try:
        # Switch back to main content
        driver.switch_to.default_content()

        # Log all visible buttons to confirm presence
        buttons = driver.find_elements(By.TAG_NAME, "button")
        print("🔘 Buttons visible on page:")
        for b in buttons:
            print("   ➤", b.text.strip())

        export_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Export Data')]"))
        )
        export_btn.click()
        print("📦 Clicked 'Export Data' button.")
    except Exception as e:
        print(f"❌ Failed to click 'Export Data' button: {e}")




print("⏳ Waiting for download to start...")
download_complete = False

while not download_complete:
    time.sleep(1)
    partial_files = glob.glob(os.path.join(download_dir, "*.crdownload"))
    full_files = glob.glob(os.path.join(download_dir, "*.csv")) + glob.glob(os.path.join(download_dir, "*.xlsx"))

    if partial_files:
        print("⬇️  Download in progress...")
    elif full_files:
        print(f"✅ Download complete: {os.path.basename(full_files[0])}")
        download_complete = True
    else:
        print("❌ No download detected yet. Still waiting...")

    
driver.quit()
