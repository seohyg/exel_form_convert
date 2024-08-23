import pandas as pd
import os
import numpy as np
from io import StringIO
from datetime import datetime

def xlsx_to_utf8_csv_in_memory(xlsx_file):
    try:
        df = pd.read_excel(xlsx_file)
        csv_buffer = StringIO()
        df.to_csv(csv_buffer, index=False, encoding='utf-8')
        csv_buffer.seek(0)
        print("Excel 파일이 메모리 상에서 UTF-8 인코딩의 CSV 데이터로 변환되었습니다.")
        return csv_buffer
    except Exception as e:
        print(f"Excel 파일 변환 중 오류 발생: {str(e)}")
        return None

def read_csv_data(csv_data):
    try:
        # 먼저 헤더 없이 데이터를 읽습니다
        csv_data.seek(0)  # 스트림 위치를 처음으로 초기화
        df = pd.read_csv(csv_data, encoding='utf-8', header=None)
        
        # 헤더를 찾기 위해 처음 몇 행을 검사합니다
        header_row = None
        for i in range(min(10, len(df))):  # 처음 10행 또는 전체 행 중 더 작은 수만큼 검사
            if '품목명' in df.iloc[i].values:  
                header_row = i
                break
        
        if header_row is not None:
            # 헤더를 찾았다면, 해당 행을 헤더로 설정하고 데이터를 다시 읽습니다
            csv_data.seek(0)  # 스트림 위치를 다시 초기화
            df = pd.read_csv(csv_data, encoding='utf-8', header=header_row)
            print(f"\n헤더를 {header_row+1}번째 행에서 찾았습니다.")
        else:
            print("헤더를 찾지 못했습니다. 첫 번째 행을 헤더로 사용합니다.")
            csv_data.seek(0)  # 스트림 위치를 다시 초기화
            df = pd.read_csv(csv_data, encoding='utf-8')
        
        print("\n파일 헤더:")
        print(df.columns.tolist())
        return df
    except Exception as e:
        print(f"CSV 데이터를 읽는 중 오류 발생: {str(e)}")
        return None
    
def process_data(df):
    new_columns = [
        '1차 카테고리', '2차 카테고리', '3차 카테고리', '상품코드', '공급사코드', 
        '묶음코드', '상품명', 'CasNO', '소비자가', '판매가격', '브랜드', '매입처', '단위', 
        '노출여부', '판매여부', '매입가', '규격', '메인이미지', '상세페이지'
    ]
    
    column_mapping = {
        '공급사코드': '품목코드',
        '상품명': '품목명',
        '소비자가': '소비자가\n(포함가)',
        '판매가격': '소비자가\n(포함가)',
        '매입가': '서주가격\n(포함가)',
        '단위': '단위',
        '규격': '규격',
        'CasNO': 'CasNO'
    }
    
    new_df = pd.DataFrame(columns=new_columns)
    
    for new_col, old_col in column_mapping.items():
        if old_col in df.columns:
            new_df[new_col] = df[old_col]
    
    new_df['브랜드'] = ''
    new_df['매입처'] = '우양메디칼'
    new_df['노출여부'] = '노출'
    new_df['판매여부'] = '판매'
    
    return new_df

def save_to_excel(df, output_file):
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            worksheet = writer.sheets['Sheet1']
            for idx, col in enumerate(df.columns, 1):
                worksheet.column_dimensions[chr(64 + idx)].width = 15
        print(f"변환된 파일이 {output_file}에 저장되었습니다.")
    except Exception as e:
        print(f"파일 저장 중 오류 발생: {str(e)}")

def convert_file_format(input_file, output_file):
    _, file_extension = os.path.splitext(input_file)
    
    if file_extension.lower() == '.xlsx':
        csv_data = xlsx_to_utf8_csv_in_memory(input_file)
        if csv_data:
            df = read_csv_data(csv_data)
        else:
            return
    elif file_extension.lower() == '.csv':
        with open(input_file, 'r', encoding='utf-8') as file:
            csv_data = StringIO(file.read())
        df = read_csv_data(csv_data)
    else:
        print("지원하지 않는 파일 형식입니다. .xlsx 또는 .csv 파일만 지원합니다.")
        return
    
    if df is not None:
        processed_data = process_data(df)
        save_to_excel(processed_data, output_file)

def main():
    input_file = input("입력 파일 경로를 입력하세요 (.xlsx 또는 .csv): ")
    
    # 현재 날짜를 YYYYMMDD 형식으로 가져오기
    current_date = datetime.now().strftime("%Y%m%d")
    
    # 출력 파일명 생성
    output_file = f"{current_date}_bot_나비엠알오_형식화.xlsx"
    
    # 출력 파일의 전체 경로 생성
    output_path = os.path.join(os.path.dirname(input_file), output_file)

    if os.path.exists(input_file):
        convert_file_format(input_file, output_path)
    else:
        print("입력한 파일이 존재하지 않습니다.")
        
if __name__ == "__main__":
    main()
