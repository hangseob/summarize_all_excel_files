import os
import shutil
from glob import glob
from datetime import datetime

def copy_files_with_keyword(source_dir, target_dir, keyword):
    """
    source_dir 내의 모든 서브폴더를 검색하여 
    파일명에 keyword가 포함된 파일만 target_dir로 복사합니다.
    """
    
    # 타겟 폴더가 없으면 생성
    if not os.path.exists(target_dir):
        os.makedirs(target_dir)
        print(f"Created target directory: {target_dir}")

    # 재귀적으로 모든 파일 검색
    search_pattern = os.path.join(source_dir, "**", f"*{keyword}*")
    # 파일만 필터링하여 리스트 생성
    files_to_copy = [f for f in glob(search_pattern, recursive=True) if os.path.isfile(f)]
    total_files = len(files_to_copy)

    if total_files == 0:
        print(f"No files found with keyword '{keyword}' in {source_dir}")
        return

    print(f"Found {total_files} files to copy. Starting...")

    count = 0
    for idx, file_path in enumerate(files_to_copy, 1):
        filename = os.path.basename(file_path)
        name, ext = os.path.splitext(filename)
        
        dest_path = os.path.join(target_dir, filename)
        
        # 파일명이 겹칠 경우 처리
        if os.path.exists(dest_path):
            parent_name = os.path.basename(os.path.dirname(file_path))
            new_filename = f"{name}_{parent_name}{ext}"
            dest_path = os.path.join(target_dir, new_filename)
            
            if os.path.exists(dest_path):
                timestamp = datetime.now().strftime("%H%M%S")
                dest_path = os.path.join(target_dir, f"{name}_{parent_name}_{timestamp}{ext}")
        
        try:
            shutil.copy2(file_path, dest_path)
            print(f"[{idx}/{total_files}] Copied: {os.path.basename(dest_path)}")
            count += 1
        except Exception as e:
            print(f"[{idx}/{total_files}] Error copying {filename}: {e}")

    print("-" * 30)
    print(f"Total {count} files copied to '{target_dir}'.")

if __name__ == "__main__":
    # 설정값 예시
    # SOURCE에 "Y:\" 같은 드라이브 명을 직접 넣어도 잘 작동합니다.
    # r"Y:\" 처럼 끝에 백슬래시가 하나만 있으면 파이썬에서 문법 에러가 발생하므로 
    # "Y:\\" 또는 "Y:/" 형식을 권장합니다.
    SOURCE = "sample_folder"  
    TARGET = "collected_data" 
    KEYWORD = "손익결과"

    copy_files_with_keyword(SOURCE, TARGET, KEYWORD)
