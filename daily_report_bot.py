import os
import time
import subprocess
from datetime import datetime

import pyautogui
import pyperclip
from openpyxl import Workbook, load_workbook

# Safety settings
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1

# Create today's date and time automatically
now = datetime.now()
date_time = now.strftime("%Y-%m-%d %H:%M:%S")
today = now.strftime("%Y-%m-%d")

# File names
excel_file = os.path.abspath(f"daily_report_{today}.xlsx")
screenshot_file = os.path.abspath(
    f"daily_report_screenshot_{today}.png"
)

# Step 1: Open BBC News in Chrome
pyautogui.hotkey("win", "r")
time.sleep(2)

pyautogui.write("https://www.bbc.com/news", interval=0.03)
pyautogui.press("enter")

# Wait for the website to load
time.sleep(12)

# Step 2: Copy webpage text
pyautogui.click(500, 350)
pyautogui.hotkey("ctrl", "a")
pyautogui.hotkey("ctrl", "c")
time.sleep(2)

page_text = pyperclip.paste()

# Extract a useful-looking line from the copied text
lines = [
    line.strip()
    for line in page_text.splitlines()
    if len(line.strip()) > 25
]

headline = lines[0] if lines else "Could not read BBC News text"
headline = headline[:300]

print("Copied information:", headline)

# Step 3: Create the Excel report
if os.path.exists(excel_file):
    workbook = load_workbook(excel_file)
    sheet = workbook.active
else:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Daily News Report"
    sheet.append(["Date & Time", "News Information", "Comment"])

sheet.append([
    date_time,
    headline,
    "Daily BBC News update"
])

sheet.column_dimensions["A"].width = 22
sheet.column_dimensions["B"].width = 65
sheet.column_dimensions["C"].width = 30

for cell in sheet[1]:
    cell.font = cell.font.copy(bold=True)

workbook.save(excel_file)

# Step 4: Open the saved Excel file
os.startfile(excel_file)
time.sleep(5)

# Step 5: Take a screenshot
pyautogui.screenshot().save(screenshot_file)

print("Excel report saved:", excel_file)
print("Screenshot saved:", screenshot_file)
print("Assignment completed!")
