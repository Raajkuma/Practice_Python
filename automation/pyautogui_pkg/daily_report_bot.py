import subprocess
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import pyautogui
import pyperclip


# ============================================================
# CONFIGURATION
# ============================================================

CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

EXCEL_PATH = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"


# Stocks to fetch
STOCKS = [
    "ITC",
    "Wipro",
    "TCS",
    "Tata Motors"
]


# ============================================================
# CURRENT DIRECTORY / FILE NAMES
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

TODAY = datetime.now()

DATE = TODAY.strftime("%Y-%m-%d")

DATE_TIME = TODAY.strftime("%d-%m-%Y %H:%M")

REPORT_FILE = BASE_DIR / f"daily_report_{DATE}.xlsx"

SCREENSHOT_FILE = BASE_DIR / f"daily_report_{DATE}.png"

BROWSER_SCREENSHOT = BASE_DIR / "stock_prices_google_page.png"


# ============================================================
# HELPER
# ============================================================

def wait(seconds=2):
    time.sleep(seconds)


# ============================================================
# OPEN CHROME
# ============================================================

def open_chrome():

    print("Opening Chrome...")

    subprocess.Popen([
        CHROME_PATH,
        "--new-window",
        "https://www.google.com"
    ])

    wait(6)

    print("Chrome opened.")


# ============================================================
# FETCH STOCK PRICE
# ============================================================

def fetch_stock_price(stock_name):

    print()
    print("------------------------------------------")
    print(f"Fetching {stock_name} stock price...")
    print("------------------------------------------")

    # Make Chrome active
    pyautogui.click(800, 400)

    wait(1)

    # Google search URL
    search_url = (
        "https://www.google.com/search?q="
        + stock_name.replace(" ", "+")
        + "+share+price"
    )

    # Open search
    pyautogui.hotkey("ctrl", "l")

    wait(1)

    pyautogui.write(
        search_url,
        interval=0.01
    )

    pyautogui.press("enter")

    wait(3)
    # Scroll slightly
    pyautogui.scroll(-500)

    wait(2)

    # Save screenshot for debugging
    pyautogui.screenshot(
        str(BROWSER_SCREENSHOT)
    )

    print(
        f"Browser screenshot saved: {BROWSER_SCREENSHOT}"
    )

    # --------------------------------------------------------
    # Select stock price
    # --------------------------------------------------------
    #
    # IMPORTANT:
    # This coordinate must point to the stock price
    # displayed by Google.
    #
    # Change it if the price appears somewhere else.
    #
    # --------------------------------------------------------

    print(f"Selecting {stock_name} price...")

    pyautogui.doubleClick(
        370,
        400
    )

    wait(1)

    # Copy selected value
    pyautogui.hotkey(
        "ctrl",
        "c"
    )

    wait(1)

    # Read clipboard
    stock_price = pyperclip.paste().strip()

    print(
        f"{stock_name} copied value: {stock_price}"
    )

    if not stock_price:

        print(
            f"WARNING: {stock_name} price was not copied."
        )

        stock_price = "Price not copied"

    return stock_price


# ============================================================
# FETCH ALL STOCKS
# ============================================================

def fetch_all_stocks():

    stock_data = []

    for stock in STOCKS:

        price = fetch_stock_price(stock)

        stock_data.append({
            "Date & Time": DATE_TIME,
            "Stock": stock,
            "Stock Price": price
        })

        # Small delay before next stock
        wait(2)

    return stock_data


# ============================================================
# CREATE PANDAS DATAFRAME
# ============================================================

def create_dataframe(stock_data):

    print()
    print("Creating Pandas DataFrame...")

    df = pd.DataFrame(stock_data)

    print()
    print("==========================================")
    print("STOCK DATA")
    print("==========================================")
    print(df)
    print()

    return df


# ============================================================
# SAVE DATAFRAME TO EXCEL
# ============================================================

def save_excel(df):

    print("Saving DataFrame to Excel...")

    # If today's report already exists, remove it
    if REPORT_FILE.exists():

        try:

            REPORT_FILE.unlink()

            print("Existing report removed.")

        except PermissionError:

            raise PermissionError(
                f"Please close Excel before running the script:\n"
                f"{REPORT_FILE}"
            )

    # Save DataFrame
    df.to_excel(
        REPORT_FILE,
        index=False,
        engine="openpyxl"
    )

    wait(2)

    # Verify file
    if REPORT_FILE.exists():

        print()
        print("==========================================")
        print("Excel file saved successfully!")
        print("==========================================")
        print(f"Location:")
        print(REPORT_FILE)
        print()

    else:

        raise FileNotFoundError(
            f"Excel file was not created: {REPORT_FILE}"
        )


# ============================================================
# OPEN SAVED EXCEL FILE
# ============================================================

def open_excel():

    print("Opening saved Excel file...")

    subprocess.Popen([
        EXCEL_PATH,
        str(REPORT_FILE)
    ])

    wait(7)

    print("Excel file opened.")


# ============================================================
# TAKE SCREENSHOT
# ============================================================

def take_screenshot():

    print("Taking screenshot of Excel...")

    # Make sure Excel is active
    pyautogui.click(
        800,
        400
    )

    wait(2)

    screenshot = pyautogui.screenshot()

    screenshot.save(
        str(SCREENSHOT_FILE)
    )

    print()
    print("==========================================")
    print("Screenshot saved successfully!")
    print("==========================================")
    print(f"Location:")
    print(SCREENSHOT_FILE)
    print()


# ============================================================
# CLOSE EXCEL
# ============================================================

def close_excel():

    print("Closing Excel...")

    pyautogui.hotkey(
        "alt",
        "f4"
    )

    wait(3)

    print("Excel closed.")


# ============================================================
# CLOSE CHROME
# ============================================================

def close_chrome():

    print("Closing Chrome...")

    # Switch to Chrome
    pyautogui.hotkey(
        "alt",
        "tab"
    )

    wait(1)

    # Close Chrome
    # pyautogui.hotkey(
    #     "alt",
    #     "f4"
    # )

    wait(3)

    print("Chrome closed.")


# ============================================================
# MAIN
# ============================================================

def main():

    print()
    print("==============================================")
    print(" DAILY STOCK STATUS REPORT AUTOMATION")
    print("==============================================")
    print()

    try:

        # ----------------------------------------------------
        # 1. Open Chrome
        # ----------------------------------------------------

        open_chrome()

        # ----------------------------------------------------
        # 2. Fetch all stocks
        # ----------------------------------------------------

        stock_data = fetch_all_stocks()

        # ----------------------------------------------------
        # 3. Create DataFrame
        # ----------------------------------------------------

        df = create_dataframe(
            stock_data
        )

        # ----------------------------------------------------
        # 4. Save DataFrame as Excel
        # ----------------------------------------------------

        save_excel(df)

        # ----------------------------------------------------
        # 5. Open saved Excel file
        # ----------------------------------------------------

        open_excel()

        # ----------------------------------------------------
        # 6. Take screenshot
        # ----------------------------------------------------

        take_screenshot()

        print()
        print("==============================================")
        print(" AUTOMATION COMPLETED SUCCESSFULLY")
        print("==============================================")
        print()

    finally:

        # ----------------------------------------------------
        # 7. Close Excel
        # ----------------------------------------------------

        close_excel()

        # ----------------------------------------------------
        # 8. Close Chrome
        # ----------------------------------------------------

        close_chrome()

        print()
        print("All applications closed.")
        print()


# ============================================================
# PROGRAM START
# ============================================================

if __name__ == "__main__":
    main()