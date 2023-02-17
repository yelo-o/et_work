from openpyxl import load_workbook

# 엑셀 파일 ws
wb = load_workbook(filename='22년 4분기 12월 매입매출정산서(비삼성)_샘플.xlsx')

# 워크시트 선택
ws = wb.active

# A열의 값을 가져오기
list0 = []
list1 = []
com_list = []

for cell in ws['A']:
    list0.append(cell.value)

print('list0은',list0)
print(len(list0))

for value in list0:
    if value not in list1:
        list1.append(value)


print('list1은',list1)
print(len(list1))
com_list = list1
del com_list[0:2]

print('최종',com_list)
print(len(com_list))

for com in com_list:

    list_to_delete = []  # 삭제해야할 행의 인덱스를 저장할 리스트 [0, 1, 2, ...]
    for row in ws.iter_rows(min_row=2, min_col=1):
        list_to_delete.append(row[0].row)
        
    list_to_delete.reverse()  # 하나씩 append했기 때문에 순서가 뒤바뀜
        
    for row_index in list_to_delete:
        ws.delete_rows(row_index)
        wb.save(f'{com}.xlsx')

    # print(list_to_delete)
    
# 변경된 내용 저장

# # 삭제할 행의 인덱스를 저장할 리스트
# # 행 반복
# for row in ws.iter_rows(min_row=2, min_col=1, values_only=True):
#     # A열이 '현대중공업'이거나 '#공통'이 아닌 경우
#     if row[0] != '현대중공업' and row[0] != '#공통':
#         # 삭제할 행의 인덱스를 리스트에 추가
#         delete_rows.append(row[0].row)

# # 삭제할 행을 역순으로 정렬하여, 뒤에서부터 삭제하도록 함
# delete_rows.reverse()

# # 행 ws
# for row in delete_rows:
#     ws.delete_rows(row)
