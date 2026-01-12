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
        
        # 대상 시트 찾기 로직 수정
        sheet = None
        # 1. 시트 이름 우선 순위: "신양식_NEW"
        for s in wb.sheets:
            if s.name == "신양식_NEW":
                sheet = s
                break
        
        # 2. 첫 번째 행에 "일일 운용 보고서" 문구가 포함된 시트 찾기
        if not sheet:
            for s in wb.sheets:
                try:
                    # 첫 행(1행)의 값을 가져옴 (보통 A1:Z1 범위 확인)
                    first_row = s.range("1:1").value
                    if first_row:
                        # 리스트 내에 "일일 운용 보고서" 문구가 포함된 셀이 있는지 확인
                        if any("일일 운용 보고서" in str(cell) for cell in first_row if cell):
                            sheet = s
                            break
                except:
                    continue
        
        # 3. 기존 Fallback: "구분"이 있는 시트 검색
        if not sheet:
            for s in wb.sheets:
                try:
                    v = s.used_range.value
                    if v:
                        # 띄어쓰기 제거 후 "구분" 검색
                        flat_top = [str(c).replace(" ", "") for row in v[:20] for c in row if c]
                        if "구분" in flat_top:
                            sheet = s
                            break
                except: continue
        
        if not sheet:
            sheet = wb.sheets[0]

        print(f"Using sheet: {sheet.name}")
        vals = sheet.used_range.value
        
        if not vals:
            print(f"Sheet {sheet.name} is empty.")
            return []

        # 섹션 헤더 및 컬럼 헤더 위치 찾기
        # 맨 위행부터 내려가면서 처음으로 "손익"을 포함한 셀 찾기
        pnl_header_pos = None # (r, c)
        pos_header_pos = None # (r, c)
        
        for r_idx, row in enumerate(vals):
            if row:
                for c_idx, cell in enumerate(row):
                    # "손익"이라는 정확한 혹은 포함된 텍스트 찾기
                    if cell and "손익" in str(cell) and len(str(cell).strip()) <= 10:
                        pnl_header_pos = (r_idx, c_idx)
                        # 같은 행의 오른편에서 "Position" 찾기
                        for c_idx_right in range(c_idx + 1, len(row)):
                            if row[c_idx_right] and ("Position" in str(row[c_idx_right]) or "포지션" in str(row[c_idx_right])):
                                pos_header_pos = (r_idx, c_idx_right)
                                break
                        if pnl_header_pos: break
            if pnl_header_pos: break
                
        if not pnl_header_pos:
            print(f"Could not find '손익' section header in {file_path}")
            return []

        # 섹션별 작업 정의 (헤더 위치, 섹션 타입)
        sections_to_process = [
            (pnl_header_pos, "손익"),
        ]
        if pos_header_pos:
            sections_to_process.append((pos_header_pos, "포지션"))

        for (h_r, h_c), s_type in sections_to_process:
            # 컬럼 헤더는 섹션 헤더의 두 행 아래 (r + 2)
            col_header_r = h_r + 2
            if col_header_r >= len(vals): continue
            
            # "구분" 셀 위치 찾기 (섹션 헤더 열 h_c 주변에서 검색)
            name_col_idx = None
            col_header_row = vals[col_header_r]
            
            # h_c 근처에서 "구분" 찾기 (띄어쓰기 포함 대응)
            for c_search in range(max(0, h_c - 2), min(len(col_header_row), h_c + 5)):
                cell_val = str(col_header_row[c_search]).replace(" ", "") if col_header_row[c_search] else ""
                if cell_val == "구분":
                    name_col_idx = c_search
                    break
            
            if name_col_idx is None:
                print(f"  - Could not find '구분' for section {s_type} at row {col_header_r}")
                continue
                
            # 값(Value) 컬럼 찾기
            val_col_idx = name_col_idx + 1 # 기본값
            if s_type == "손익":
                for c_search in range(name_col_idx + 1, len(col_header_row)):
                    clean_header = str(col_header_row[c_search]).replace(" ", "") if col_header_row[c_search] else ""
                    if clean_header == "일":
                        val_col_idx = c_search
                        break
            else:
                for c_search in range(name_col_idx + 1, len(col_header_row)):
                    if col_header_row[c_search] and "부서" in str(col_header_row[c_search]):
                        val_col_idx = c_search
                        break
            
            print(f"  - Section [{s_type}] found. Header Row: {h_r}, Name Col: {name_col_idx}, Value Col: {val_col_idx}")

            section_sum_calc = 0
            section_total_excel = None
            
            curr_r = col_header_r + 1
            while curr_r < len(vals):
                name = vals[curr_r][name_col_idx]
                value = vals[curr_r][val_col_idx]
                
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
                        "손익/포지션": s_type,
                        "구분": name,
                        "값": f_val,
                        "통화": currency,
                        "단위": unit
                    })
                    if s_type == "손익":
                        section_sum_calc += f_val
                curr_r += 1

            if s_type == "손익" and section_total_excel is not None:
                diff = abs(section_sum_calc - section_total_excel)
                if diff <= 2:
                    print(f"  [정합성 체크] {s_type}: 정상 (합계: {section_total_excel})")
                else:
                    print(f"  [정합성 체크] {s_type}: 불일치! (계산: {section_sum_calc}, 엑셀: {section_total_excel}, 차이: {diff})")

        # "조달금리 감안 총캐리" 추가 추출

        # "조달금리 감안 총캐리" 추가 추출
        for r_idx, row in enumerate(vals):
            for c_idx, cell_val in enumerate(row):
                if cell_val and ("조달금리" in str(cell_val) and "총캐리" in str(cell_val)):
                    val_str = str(cell_val)
                    # 1. 같은 셀 내에 숫자(및 마이너스 부호)가 있는지 확인
                    match = re.search(r'[:\s]+([-?\d.]+)', val_str)
                    if match:
                        try:
                            val = float(match.group(1).replace(',', ''))
                            results.append({
                                "date": date_str, "손익/포지션": "손익_합산제외", "구분": "조달금리 감안 총캐리",
                                "값": val, "통화": "KRW", "단위": unit
                            })
                            print(f"  - Found Carry P&L in cell: {val}")
                        except: pass
                    else:
                        # 2. 같은 셀에 없으면 오른쪽 셀들 확인
                        for offset in range(1, 5):
                            if c_idx + offset < len(row):
                                f_val = to_float(row[c_idx + offset])
                                if f_val is not None and f_val != 0:
                                    results.append({
                                        "date": date_str, "손익/포지션": "손익_합산제외", "구분": "조달금리 감안 총캐리",
                                        "값": f_val, "통화": "KRW", "단위": unit
                                    })
                                    print(f"  - Found Carry P&L in neighbor cell: {f_val}")
                                    break
                    break

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
