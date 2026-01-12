import os
import pandas as pd
import xlwings as xw
import re
from glob import glob
import time
from datetime import datetime

def process_excel(app, file_path):
    print(f"Processing {file_path}...")
    
    # Extract date from filename: 손익결과_YYYYMMDD.xlsm
    filename = os.path.basename(file_path)
    date_match = re.search(r'(\d{8})', filename)
    if not date_match:
        date_match = re.search(r'(\d{8})', file_path)
    
    date_str = date_match.group(1) if date_match else "Unknown"
    
    wb = None
    results = []
    unit = 1000000
    
    def to_float(val):
        if val is None:
            return 0.0
        if isinstance(val, (int, float)):
            return float(val)
        try:
            # Handle cases like "-" which Excel often uses for 0
            s = str(val).strip().replace(',', '')
            if s == '-':
                return 0.0
            return float(s)
        except:
            return None

    try:
        wb = app.books.open(file_path, read_only=True, update_links=False)
        time.sleep(2) # Wait for DRM/Macros
        
        # Find sheet containing "당일"
        sheet = None
        for s in wb.sheets:
            if "당일" in s.name:
                sheet = s
                break
        
        if not sheet:
            # Fallback: look for "구분" in any sheet
            for s in wb.sheets:
                try:
                    v = s.used_range.value
                    if v:
                        # Flatten first 20 rows to look for "구분"
                        flat_top = [str(c) for row in v[:20] for c in row if c]
                        if "구분" in flat_top:
                            sheet = s
                            break
                except:
                    continue
        
        if not sheet:
            # Last resort: first sheet
            sheet = wb.sheets[0]

        print(f"Using sheet: {sheet.name}")
        vals = sheet.used_range.value
        
        if not vals:
            print(f"Sheet {sheet.name} is empty.")
            return []

        # Find "손익" and "Position/Risk" sections
        for r_idx, row in enumerate(vals):
            for c_idx, cell_val in enumerate(row):
                if cell_val == "구분":
                    section_header = ""
                    if r_idx > 0:
                        for offset in range(1, 4):
                            if r_idx - offset >= 0:
                                v = vals[r_idx - offset][c_idx]
                                if v:
                                    section_header = str(v)
                                    break
                    
                    section_type = "손익" if "손익" in section_header else "포지션"
                    
                    section_sum_calc = 0
                    section_total_excel = None
                    
                    curr_r = r_idx + 1
                    while curr_r < len(vals):
                        name = vals[curr_r][c_idx]
                        value = vals[curr_r][c_idx + 1]
                        
                        # Handle "합계" row
                        if name and str(name).strip() == "합계":
                            section_total_excel = to_float(value)
                            break
                            
                        if name is None or str(name).strip() == "":
                            break
                        
                        f_val = to_float(value)
                        if f_val is not None:
                            currency = "USD" if "$" in str(name) else "KRW"
                            results.append({
                                "date": date_str,
                                "손익/포지션": section_type,
                                "구분": name,
                                "값": f_val,
                                "통화": currency,
                                "단위": unit
                            })
                            if section_type == "손익":
                                section_sum_calc += f_val
                        curr_r += 1

                    # 정합성 체크 (손익 섹션에 대해서만)
                    if section_type == "손익" and section_total_excel is not None:
                        diff = abs(section_sum_calc - section_total_excel)
                        # 절사 고려하여 차이가 2 이하인 경우 정상으로 판단
                        if diff <= 2:
                            print(f"  [정합성 체크] {section_type}: 정상 (합계: {section_total_excel})")
                        else:
                            print(f"  [정합성 체크] {section_type}: 불일치! (계산: {section_sum_calc}, 엑셀: {section_total_excel}, 차이: {diff})")

        return results

    except Exception as e:
        print(f"Error processing {file_path}: {e}")
        return []
    finally:
        if wb:
            try:
                wb.close()
            except:
                pass

def main():
    target_pattern = os.path.join("sample_folder", "**", "손익결과*.xlsm")
    files = glob(target_pattern, recursive=True)
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    
    if not files:
        print("No matching files found.")
        return

    app = xw.App(visible=True, add_book=False)
    
    all_data = []
    try:
        for f in files:
            all_data.extend(process_excel(app, f))
    finally:
        app.quit()
    
    if not all_data:
        print("No data extracted.")
        return
        
    df = pd.DataFrame(all_data)
    cols = ["date", "손익/포지션", "구분", "값", "통화", "단위"]
    for col in cols:
        if col not in df.columns:
            df[col] = None
    df = df[cols]
    
    # 파일명 중복 방지를 위한 타임스탬프 추가 (YYMMDD_HHMMSS)
    timestamp = datetime.now().strftime("%y%m%d_%H%M%S")
    output_file = f"취합결과_{timestamp}.xlsx"
    
    df.to_excel(output_file, index=False)
    print(f"Saved results to {output_file}")

if __name__ == "__main__":
    main()
