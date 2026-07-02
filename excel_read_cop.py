import pandas as pd

# Excel 파일 경로
file_path = "example.xlsx"

# Excel 파일 읽기
try:
    df = pd.read_excel(file_path, engine='openpyxl')  # openpyxl 엔진 사용
    print("Excel 파일 읽기 성공!")
    print(df.head())  # 데이터의 첫 5줄 출력
except FileNotFoundError:
    print(f"파일을 찾을 수 없습니다: {file_path}")
except Exception as e:
    print(f"오류 발생: {e}")