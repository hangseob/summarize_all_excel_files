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
    # ** 는 모든 서브폴더를 의미하며, recursive=True와 함께 사용해야 함
    search_pattern = os.path.join(source_dir, "**", f"*{keyword}*")
    files_to_copy = glob(search_pattern, recursive=True)

    count = 0
    for file_path in files_to_copy:
        # 디렉토리는 제외하고 파일만 처리
        if os.path.isfile(file_path):
            filename = os.path.basename(file_path)
            name, ext = os.path.splitext(filename)
            
            dest_path = os.path.join(target_dir, filename)
            
            # 파일명이 겹칠 경우 처리
            if os.path.exists(dest_path):
                # 파일 경로의 일부를 이름 뒤에 붙여 중복 방지
                # 부모 폴더명을 가져와서 이름 뒤에 붙임
                parent_name = os.path.basename(os.path.dirname(file_path))
                new_filename = f"{name}_{parent_name}{ext}"
                dest_path = os.path.join(target_dir, new_filename)
                
                # 만약 부모 폴더명을 붙였는데도 겹치면 타임스탬프 등을 추가로 붙임
                if os.path.exists(dest_path):
                    timestamp = datetime.now().strftime("%H%M%S")
                    dest_path = os.path.join(target_dir, f"{name}_{parent_name}_{timestamp}{ext}")
            
            try:
                shutil.copy2(file_path, dest_path)
                print(f"Copied: {os.path.basename(dest_path)} from {os.path.dirname(file_path)}")
                count += 1
            except Exception as e:
                print(f"Error copying {filename}: {e}")

    print("-" * 30)
    print(f"Total {count} files copied to '{target_dir}'.")

if __name__ == "__main__":
    # 설정값 예시
    # SOURCE에 "Y:\" 같은 드라이브 명을 직접 넣어도 잘 작동합니다.
    # SOURCE = r"Y:\"  
    SOURCE = "sample_folder"
    
    # r을 붙여 raw string으로 처리하는 것이 안전합니다.
    TARGET = "collected_data" # 폴더가 없으면 자동으로 생성됩니다.
    KEYWORD = "손익결과"

    copy_files_with_keyword(SOURCE, TARGET, KEYWORD)
