# pdf 이미지 변환 부분
from pdf2image import convert_from_path
import os

# 모듈 import
import re
import pyautogui as pg
# gui 부분
import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import *
from tkinter import filedialog


pop_path = 'C:/flyordig/et_work/poppler-23.01.0/Library/bin'

def file_select_btn():
    global pdf_path, pdf
    root.filename = filedialog.askopenfilename(title = '변환할 PDF 파일을 골라주세요.',
                                           filetypes = (("PDF","*.pdf"),("all files","*.*")))
    pdf_path = root.filename
    print(pdf_path)
    pdf = re.search(r'([^\\/]+$)', pdf_path).group(1)
    # pdf_path = "C:/flyordig/et_work/paper.pdf"
    selected_files.configure(text="{}".format(pdf))

def change_to_jpg():
    global pop_path
     
    pages = convert_from_path(pdf_path, poppler_path=pop_path)

    for i, page in enumerate(pages):
        filename = os.path.splitext(os.path.basename(pdf))[0] + f"_{i+1}페이지.jpg"
        page.save(filename, 'JPEG')
    pg.alert("jpg 파일로 변환이 완료되었습니다.")



# GUI 부분
root = ttk.Window()
root.title("pdf 합치기 도구")
frame1 = ttk.Frame(root, width=700, height=500, bootstyle = "default") # 1번 프레임 좌측에 배치
frame1.pack(fill=tk.X, expand=True)

# 프레임 1의 첫번째 열
first_label = Label(frame1, text = "PDF 파일을 선택해주세요.")
btn_search = ttk.Button(frame1, text="PDF 파일 불러오기", bootstyle="info", command=file_select_btn)
# selected_file_list = Listbox(frame1)
selected_files = ttk.Label(frame1, text="", bootstyle="inverse-dark")
btn_merge = ttk.Button(frame1, text="jpg로 변환하기", bootstyle="secondary", command=change_to_jpg)

# grid 정렬
first_label.grid(row=1, column=0, sticky="nsew",padx=5, pady=5)
btn_search.grid(row=3, column=0, sticky="ew", padx=5, pady=5)
# selected_file_list.grid(row=3, column=1, sticky="ns",padx=5, pady=5)
selected_files.grid(row=3, column=1, sticky="ns",padx=5, pady=5)
btn_merge.grid(row=3, column=2, sticky="ew", padx=5, pady=5 )


root.mainloop()
