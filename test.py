import pandas as pd
from openpyxl import load_workbook

df = pd.read_excel('C:/flyordig/et_work/22년 4분기 12월 매입매출정산서(비삼성)_샘플.xlsx', sheet_name=2, engine='openpyxl')
# df = pd.read_excel('C:/flyordig/et_work/22년 4분기 12월 매입매출정산서(비삼성)_샘플.xlsx', sheet_name=2, engine='openpyxl', read_only=True)



df = df[df.회사 == '현대글로벌서비스']

print("파일 저장")

with pd.ExcelWriter('example.xlsx', mode='a', engine='openpyxl') as writer:
    writer.book = load_workbook('example.xlsx')
    df.to_excel(writer, sheet_name='3번시트', index=False)