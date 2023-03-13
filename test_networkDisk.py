import pandas as pd

# 네트워크 경로 지정
network_path = r'O:/김민규/file.xlsx'

# 엑셀 파일 불러오기
df = pd.read_excel(network_path)

# 데이터 프레임 확인
print(df.head())

