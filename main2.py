import time
import keyboard
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import tkinter as tk
from tkinter import filedialog, simpledialog

# --- Ask User for Excel File and Column ---
root = tk.Tk()
root.withdraw()  # Hide the main tkinter window

excel_path = filedialog.askopenfilename(
    title="Select Excel file",
    filetypes=[("Excel Files", "*.xlsx *.xls")]
)
if not excel_path:
    print("No file selected. Exiting.")
    exit()

column_letter = simpledialog.askstring(
    title="Column Selection",
    prompt="Enter column letter (e.g., B):"
)
if not column_letter:
    print("No column selected. Exiting.")
    exit()

# --- Load Excel ---
workbook = load_workbook(excel_path)
sheet = workbook.active

# --- Launch Browser ---
driver = webdriver.Chrome()
driver.get("https://in.tradingview.com/chart/")
wait = WebDriverWait(driver, 10)

time.sleep(5)  # Let the page load

# --- Function to search symbol ---
def search_symbol(symbol):
    try:
        body = driver.find_element(By.TAG_NAME, "body")
        body.send_keys(Keys.ESCAPE)
        time.sleep(0.5)

        search_button = wait.until(EC.element_to_be_clickable((By.ID, "header-toolbar-symbol-search")))
        search_button.click()

        input_field = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//input[@data-role='search' and @role='searchbox']")
        ))

        input_field.clear()
        time.sleep(0.2)
        input_field.send_keys(symbol)
        input_field.send_keys(Keys.ENTER)

        print(f"Searched for: {symbol}")
    except TimeoutException:
        print(f"Timeout: Could not find input for {symbol}")
    except Exception as e:
        print(f"Error searching symbol {symbol}: {e}")

# --- Main Execution Loop ---
current_row = 2
max_rows = 101

print("Press Ctrl+Space to go to next symbol, Esc to abort.")

while current_row <= max_rows:
    cell = f'{column_letter.upper()}{current_row}'
    symbol = sheet[cell].value
    if not symbol:
        print(f"Row {current_row} is empty. Skipping.")
        current_row += 1
        continue

    search_symbol(symbol)

    while True:
        if keyboard.is_pressed("ctrl+space"):
            current_row += 1
            break
        if keyboard.is_pressed("esc"):
            print("Aborted by user.")
            driver.quit()
            exit()
        time.sleep(0.1)
