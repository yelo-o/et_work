# 모듈 import
import re
# pdf 부분
from PyPDF2 import PdfMerger
# gui 부분
import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import *
from tkinter import filedialog

# 변수 설정
pdfs = []

def merge_pdf():
    global pdfs

    merger = PdfMerger()

    for pdf in pdfs:
        merger.append(pdf_dir+pdf)
    
    merger.write(f"{pdf_dir}합친 파일.pdf")
    merger.close()
    

# 파일 불러오기

# 파일 선택
def file_select_btn():
    global root, pdfs, pdf_dir
    root.filename = filedialog.askopenfilenames(initialdir = "C:/flyordig/et_work/docs",
                                                title = "PDF 파일을 골라주세요",
                                                filetypes = (("PDF","*.pdf"),("all files","*.*")))
    ssfiles = root.filename
    print(ssfiles)
    refined_ssfiles = []
    # pdfs = []
    for ssfile in ssfiles:
        # global refined_ssfiles
        refined_ssfiles.append(ssfile)
    for refined_ssfile in refined_ssfiles:
        ssfile2 = re.search(r'([^\\/]+$)', refined_ssfile).group(1)
        pdfs.append(ssfile2)
        print('ssfile2는',ssfile2)
    print('refine_ssfiles는',refined_ssfiles)
    print('refine_ssfiles2는',pdfs)
    pdf_dir = re.search(r"(.*\/).*", ssfile).group(1)
    selected_files.configure(text="{}".format(pdfs))
    # print(pdf_dir)
    # print(ssfile2)
    # pdfs.append(ssfile2)
    print(pdfs)
    # root.withdraw()

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
btn_merge = ttk.Button(frame1, text="합치기", bootstyle="secondary", command=merge_pdf)

# grid 정렬
first_label.grid(row=1, column=0, sticky="nsew",padx=5, pady=5)
btn_search.grid(row=3, column=0, sticky="ew", padx=5, pady=5)
# selected_file_list.grid(row=3, column=1, sticky="ns",padx=5, pady=5)
selected_files.grid(row=3, column=1, sticky="ns",padx=5, pady=5)
btn_merge.grid(row=3, column=2, sticky="ew", padx=5, pady=5 )


root.mainloop()