import xlwings as xw
import os
import re

def insert_pnl_data():
    # 파일 경로 설정 (절대 경로로 변환하여 안전하게 접근)
    base_path = os.getcwd()
    file_path = os.path.normpath(os.path.join(base_path, "sample_folder", "손익결과_20260115.xlsm"))
    
    if not os.path.exists(file_path):
        print(f"Error: {file_path} not found.")
        return

    # 입력할 데이터 목록
    pnl_data = [
        ("채권", 1864), ("금리선물", -811), ("통화선물", -2), ("주식선물", 0),
        ("이자율스왑", -320), ("외화채권", 1), ("외화예금", 0), ("원화예금", 16),
        ("해외선물", 2), ("FX 주1)", -1), ("이자비용 주2)", -265),
        ("기타손익 주3)", -9), ("합계", 475)
    ]

    app = xw.App(visible=True, add_book=False)
    try:
        # 파일이 이미 열려 있는지 확인 후 연결
        try:
            wb = app.books[os.path.basename(file_path)]
        except:
            wb = app.books.open(file_path)
            
        sheet = None
        for s in wb.sheets:
            if s.name == "신양식_NEW":
                sheet = s
                break
        if not sheet: sheet = wb.sheets[0]

        print(f"작업 시트: {sheet.name}")

        # 1. "손익" 섹션의 시작 위치 찾기 (A1부터 넓게 검색)
        search_range = sheet.range("A1:Z20").value
        section_row = None
        section_col = None
        
        for r_idx, row in enumerate(search_range):
            if row:
                for c_idx, cell in enumerate(row):
                    if cell == "손익":
                        section_row = r_idx + 1
                        section_col = c_idx + 1
                        break
            if section_row: break

        if not section_row:
            print("'손익' 제목을 찾을 수 없어 기본 위치(G열 근처)에서 검색합니다.")
            section_row, section_col = 4, 7

        # 2. "손익" 제목 아래에서 "구분"과 "일" 헤더 찾기
        # 보통 제목 2행 아래에 헤더가 있음 (6행 근처)
        header_row_idx = section_row + 2
        header_vals = sheet.range(f"{header_row_idx}:{header_row_idx}").value
        
        name_col = None
        day_col = None
        
        if header_vals:
            for c_idx, val in enumerate(header_vals):
                clean_val = str(val).replace(" ", "") if val else ""
                if clean_val == "구분" and not name_col:
                    name_col = c_idx + 1
                if clean_val == "일":
                    day_col = c_idx + 1
        
        if not name_col or not day_col:
            # 못 찾을 경우 이미지 분석 기반 강제 지정 (B열=2, G열=7, Q열=17 등 양식에 따라 다름)
            # 이미지 2를 보면 "구분"은 병합된 셀이며 데이터가 시작되는 곳은 B열(2) 또는 G열(7)로 보임
            # 여기서는 이미지상 녹색 칸의 위치인 G열(7)과 Q열(17)을 우선적으로 고려함
            name_col = name_col if name_col else 7 
            day_col = day_col if day_col else 17
            print(f"헤더 자동 탐색 실패 또는 일부 누락으로 지정: 구분({name_col}열), 일({day_col}열)")
        else:
            print(f"헤더 탐색 성공: 행 {header_row_idx}, 구분({name_col}열), 일({day_col}열)")

        # 3. 데이터 입력 (기존 오답 데이터 삭제 후 입력)
        for i, (name, val) in enumerate(pnl_data):
            row = header_row_idx + 1 + i
            # '구분' 열에 이름 입력
            sheet.range((row, name_col)).value = name
            # '일' 열에 숫자 입력
            sheet.range((row, day_col)).value = val
            print(f"입력 완료: {name} ({row}행, {name_col}열 / {day_col}열)")

        # 4. 저장 (읽기 전용 에러 방지)
        try:
            wb.save()
            print("저장 완료.")
        except:
            new_path = file_path.replace(".xlsm", "_수정본.xlsm")
            wb.save(new_path)
            print(f"읽기 전용 상태라 새 파일로 저장했습니다: {os.path.basename(new_path)}")

    except Exception as e:
        print(f"오류 발생: {e}")
    finally:
        pass

if __name__ == "__main__":
    insert_pnl_data()
