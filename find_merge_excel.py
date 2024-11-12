"""
폴더 경로를 입력받아 엑셀 파일을 찾고 하나의 엑셀 파일로 병합합니다.
"""


import os
import pandas as pd


from natsort import natsorted


def find_excel_files(folder_path):
    """ 주어진 폴더 및 서브폴더에서 엑셀 파일을 찾는 함수 """
    excel_files = []
    for root, _, files in os.walk(folder_path):
        for file in natsorted(files):
            if file.endswith(('.xlsx', '.xls')):
                excel_files.append(os.path.join(root, file))
    print(f"\n총 {len(excel_files)} 개 엑셀 파일 발견\n")
    return excel_files


def merge_excel_files(excel_files, output_path):
    """ 엑셀 파일들을 하나로 합치는 함수 """
    combined_df = pd.DataFrame()

    for file in excel_files:
        try:
            df = pd.read_excel(file)
            combined_df = pd.concat([combined_df, df], ignore_index=True)
        except Exception as e:  # pylint: disable=W0718
            print(f"파일 {file}을(를) 처리하는 동안 오류 발생: {e}")

    combined_df.to_excel(output_path, index=False)
    print(f"합쳐진 파일이 {output_path}에 저장되었습니다.")


def main():
    """main"""
    folder_path = input("엑셀 파일이 포함된 폴더 경로를 입력하세요\n=> ").strip()
    if not os.path.isdir(folder_path):
        print("유효한 폴더 경로가 아닙니다.")
        return

    output_path = input("저장할 결과 엑셀 파일 경로를 입력하세요\n=> ").strip()

    # 폴더에서 엑셀 파일 찾기
    excel_files = find_excel_files(folder_path)

    if not excel_files:
        print("해당 폴더 및 서브폴더에 엑셀 파일이 없습니다.")
        return

    # 엑셀 파일 합치기
    merge_excel_files(excel_files, output_path)


if __name__ == "__main__":
    main()
