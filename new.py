import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from datetime import datetime
from webdriver_manager.chrome import ChromeDriverManager

# ==========================
# CONFIGURATION
# ==========================
URL = "https://dc.cibil.com/DE/ccir/Login.aspx"
USERNAME = # your username
PASSWORD = # your password
EXCEL_FILE = "data.xlsx"                # must be in same folder

month_map = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
}

# ==========================
# SETUP SELENIUM
# ==========================
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# ==========================
# STEP 1: LOGIN
# ==========================
driver.get(URL)

WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "txtUsername")))
driver.find_element(By.ID, "txtUsername").send_keys(USERNAME)
driver.find_element(By.ID, "txtCred").send_keys(PASSWORD)
driver.find_element(By.ID, "logInbtn").click()

WebDriverWait(driver, 30).until_not(EC.url_contains("Login"))
print("✅ Logged in successfully")

# ==========================
# STEP 2: LOAD EXCEL DATA
# ==========================
df = pd.read_excel(EXCEL_FILE)

# ===========================
# Functions
#============================
def split_address(address, length=40):
    if pd.isna(address):  # Handle empty cells
        return ["", "", ""]
    parts = [address[i:i+length] for i in range(0, len(address), length)]
    return parts + ["", "", ""]  # Ensure at least 3 elements

# ==========================
# STEP 3: FILL FORM FOR EACH ROW
# ==========================
for index, row in df.iterrows():
    full_name = str(row['Name']).strip()
    parts = full_name.split()

    if len(parts) == 1:
        first, middle, last = parts[0], "", ""
    elif len(parts) == 2:
        first, middle, last = parts[0], "", parts[1]
    else:
        first, middle, last = parts[0], " ".join(parts[1:-1]), parts[-1]

    print(f"➡ Filling form for {first} {last}")

    # Fill name fields
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "kaf_34")))
    driver.find_element(By.NAME, "kaf_34").clear()
    driver.find_element(By.NAME, "kaf_34").send_keys(first)

    driver.find_element(By.NAME, "kaf_35").clear()
    driver.find_element(By.NAME, "kaf_35").send_keys(middle)

    driver.find_element(By.NAME, "kaf_36").clear()
    driver.find_element(By.NAME, "kaf_36").send_keys(last)

    # ==========================
    # GENDER (Dropdown Example)
    # ==========================
    gender_dropdown = Select(driver.find_element(By.NAME, "kaf_38"))  # update locator
    gender_dropdown.select_by_visible_text(row["Gender"])  # "Male" / "Female"

    # ==========================
    # STATE (Dropdown Example)
    # ==========================
    state_dropdown = Select(driver.find_element(By.NAME, "kaf_822"))  # update locator
    state_dropdown.select_by_visible_text(row["State"])  # e.g. "MH"
    time.sleep(1)
    # City
    driver.find_element(By.NAME, "kaf_113").send_keys(str(row["City"]))
    time.sleep(1)
    # Pincode
    driver.find_element(By.NAME, "kaf_114").send_keys(str(row["Pincode"]))

    # ================= DOB =================
    dob_cell = row['DOB']

    # If already datetime (Excel usually gives this)
    if isinstance(dob_cell, datetime):
        dob_value = dob_cell
    else:
    # If string like "15-08-1985"
        dob_str = str(dob_cell).strip().split(" ")[0]  # remove any time part
        try:
            dob_value = datetime.strptime(dob_str, "%d-%m-%Y")
        except ValueError:
            dob_value = datetime.strptime(dob_str, "%Y-%m-%d")

    day = dob_value.day       # 15
    month = dob_value.month   # 8
    year = dob_value.year     # 1985


    # Click on DOB field to open date picker
    dob_field = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.NAME, "kaf_37"))  # update locator if needed
    )
    dob_field.click()
    time.sleep(1)

    # Select Year
    year_dropdown = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CLASS_NAME, "ui-datepicker-year"))
    )
    Select(year_dropdown).select_by_visible_text(str(year))

    # Select Month
    month_dropdown = driver.find_element(By.CLASS_NAME, "ui-datepicker-month")
    Select(month_dropdown).select_by_index(month - 1)  # (January=0, August=7)

    # Select Day
    day_cells = driver.find_elements(By.XPATH, "//table[@class='ui-datepicker-calendar']//a")
    for d in day_cells:
        if d.text == str(day):
            d.click()
            break

    time.sleep(2)

    # ===============================
    # ADDRESS TYPE (Dropdown Example)
    # ===============================
    state_dropdown = Select(driver.find_element(By.NAME, "kaf_112"))  # update locator
    state_dropdown.select_by_visible_text("Residence Address")  # e.g. "MH"
    # ===============================
    # ADDRESS (Dropdown Example)
    # ===============================
    address_parts = split_address(str(row["Address"]))

    # Fill Address Line 1, 2, 3
    driver.find_element(By.ID, "kaf_107").send_keys(address_parts[0])
    driver.find_element(By.ID, "kaf_108").send_keys(address_parts[1])
    driver.find_element(By.ID, "kaf_109").send_keys(address_parts[2])

    # ===============================
    # MOBILE (Dropdown Example)
    # ===============================
    state_dropdown = Select(driver.find_element(By.NAME, "kaf_58"))  # update locator
    state_dropdown.select_by_visible_text("Mobile Phone")  # Default
    time.sleep(1)
    # Contact
    driver.find_element(By.NAME, "kaf_57").send_keys(str(row["Contact"]))

    # PAN
    driver.find_element(By.NAME, "kaf_40").send_keys(str(row["PAN"]))

    # Aadhaar
    driver.find_element(By.NAME, "kaf_50").send_keys(str(row["Aadhaar"]))

    #Reference Number
    driver.find_element(By.NAME, "kaf_136").send_keys(str(row["Registration"]))
    
    state_dropdown = Select(driver.find_element(By.NAME, "kaf_7"))  # update locator
    state_dropdown.select_by_visible_text("Other")  # Default

    driver.find_element(By.NAME, "kaf_9").send_keys(str("1"))



print("✅ All data submitted")
print("👉 Please verify details in the browser and press SUBMIT yourself.")
while True:
    try:
        time.sleep(60)  # Keeps script alive, browser will stay open
    except KeyboardInterrupt:
        print("❌ Script stopped manually")
        break
