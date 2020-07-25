from excel_magic.dataset import open_file
import pyautogui as auto
import time


time.sleep(10)

with open_file('财务报告.xlsx') as f:
    sheet = f.get_sheet_by_index(0)
    rows = sheet.get_rows()
    for row in rows:
        auto.keyDown("tab")
        auto.keyDown("enter")
        # find ADD
        auto.keyDown("tab")
        auto.keyDown("tab")
        auto.typewrite(row['指标名称'].value)
        auto.keyDown("tab")
        auto.typewrite(str(row['第一季度'].value))
        auto.keyDown("tab")
        auto.typewrite(str(row['第二季度'].value))
        auto.keyDown("tab")
        auto.typewrite(str(row['第三季度'].value))
        auto.keyDown("tab")
        auto.typewrite(str(row['第四季度'].value))
        auto.keyDown("tab")
        auto.keyDown("enter")
auto.keyDown("tab")
auto.keyDown("tab")
auto.keyDown("enter")
