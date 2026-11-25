import xlwings as xw
import os

def modify_excel_format():
    # 파일이 sample_folder 내부에 있으므로 경로를 수정합니다.
    file_path = os.path.join("sample_folder", "양식변환중.xlsm")
    
    if not os.path.exists(file_path):
        print(f"에러: '{file_path}' 파일을 찾을 수 없습니다.")
        return

    # 엑셀 앱 실행
    app = xw.App(visible=True, add_book=False)
    
    try:
        print(f"파일 여는 중: {file_path}")
        wb = app.books.open(file_path)
        # 작업할 시트 선택 (활성화된 시트)
        sheet = wb.sheets.active
        print(f"시트 '{sheet.name}'에서 작업을 시작합니다.")

        # 6행부터 21행까지 반복
        for row in range(6, 22):
            # 1. T열과 U열 합치기
            range_tu = sheet.range(f"T{row}:U{row}")
            val_t = sheet.range(f"T{row}").value
            val_u = sheet.range(f"U{row}").value
            
            # 내용 합치기 (None 값 제외)
            combined_tu = (str(val_t) if val_t is not None else "") + \
                          (str(val_u) if val_u is not None else "")
            
            # 셀 병합 및 값 입력
            range_tu.merge()
            range_tu.value = combined_tu
            
            # 2. V열과 W열 합치기
            range_vw = sheet.range(f"V{row}:W{row}")
            val_v = sheet.range(f"V{row}").value
            val_w = sheet.range(f"W{row}").value
            
            combined_vw = (str(val_v) if val_v is not None else "") + \
                          (str(val_w) if val_w is not None else "")
            
            # 셀 병합 및 값 입력
            range_vw.merge()
            range_vw.value = combined_vw
            
            print(f"{row}행 작업 완료 (T:U, V:W 병합 및 데이터 합침)")

        # 변경사항 저장
        wb.save()
        print("-" * 30)
        print("모든 작업이 완료되어 저장되었습니다.")

    except Exception as e:
        print(f"작업 중 오류 발생: {e}")
    finally:
        # 작업이 끝난 후 엑셀을 닫지 않고 사용자가 확인할 수 있도록 유지합니다.
        # 필요시 wb.close() 및 app.quit()를 호출할 수 있습니다.
        pass

if __name__ == "__main__":
    modify_excel_format()
