import xlwings as xw
import os

def check_value():
    file_path = os.path.join("sample_folder", "양식변환중.xlsm")
    app = xw.App(visible=False)
    try:
        wb = app.books.open(file_path)
        sheet = wb.sheets.active
        t6_val = sheet.range("T6").value
        u6_val = sheet.range("U6").value
        print(f"T6 Value: {t6_val}")
        print(f"U6 Value: {u6_val}")
        wb.close()
    finally:
        app.quit()

if __name__ == "__main__":
    check_value()
