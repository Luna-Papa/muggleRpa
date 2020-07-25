from excel_magic.dataset import Dataset

excel = Dataset("年度财务报告.xlsx")
excel.split_sheets_to_file()