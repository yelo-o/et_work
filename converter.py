# 모듈 import
import re
import os
import win32com.client

# pdf 부분
from PyPDF2 import PdfMerger
# gui 부분
import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import *
from tkinter import filedialog
# 변수 설정
ppts = []
pdfs = []


# PDF 파일 합치기
def merge_pdf():
    global pdfs
    merger = PdfMerger()
    for pdf in pdfs:
        merger.append(pdf_dir+pdf)
    
    merger.write(f"{pdf_dir}합친 파일.pdf")
    merger.close()
    
# PDF 파일 불러오기
def pdfs_select_btn():
    global root, pdfs, pdf_dir  # 글로별 변수 가져오기
    # 최초 디렉토리 미설정
    root.filename = filedialog.askopenfilenames(title = "PDF 파일을 골라주세요",
                                                filetypes = (("PDF","*.pdf"),("all files","*.*")))
    ssfiles = root.filename
    print(ssfiles)
    refined_ssfiles = []
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
    print(pdfs)


# PPT 파일 불러오기    
def ppts_select_btn():
    global root, ppts, ppt_dir
    root.filename = filedialog.askopenfilenames(title = "PDF 파일을 골라주세요",
                                                filetypes = (("PPT 파일","*.pptx"),("all files","*.*")))
    ptfiles = root.filename
    print(ptfiles)
    refined_ptfiles = []
    for ptfile in ptfiles:
        refined_ptfiles.append(ptfile)
    
    for refined_ptfile in refined_ptfiles:
        ptfile2 = re.search(r'([^\\/]+$)', refined_ptfile).group(1)
        ppts.append(ptfile2)
        ppt_dir = re.search(r"(.*\/).*", ptfile).group(1)
        selected_files2.configure(text='{}'.format(ppts))
        
    
    for ptfile in ptfiles:
        refined_ptfiles.append()
    print(ptfiles)
        

# PPT -> PDF 변환하기
def ppt_to_pdf():
    global ppts, ppt_dir
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    merger = PdfMerger()
    # for ppt in ppts:
    #     merger.append(ppt_dir+ppt)
    for ppt in ppts:
        deck = powerpoint.Presentations.Open(ppt)
        pre, ext = os.path.splitext(ppt)
        deck.SaveAs(ppt_dir+ ppt +".pdf", 32)
        deck.Close()
        # with slides.Presentation(ppt_dir+ppt) as presentation:
        #     presentation.save(f"변환완료_{ppt}.pdf", slides.export.SaveFormat.PDF)
    # merger.write(f"{ppt_dir}합친 파일.pdf")
    powerpoint.Quit()


# GUI 부분
root = ttk.Window()
root.title("pdf 합치기 도구")
frame1 = ttk.Frame(root, width=700, height=500, bootstyle = "default") # 1번 프레임 좌측에 배치
frame1.pack(fill=tk.X, expand=True)

# 프레임 1의 첫번째 열
first_label = Label(frame1, text = "PDF 파일을 선택해주세요.")
btn_search = ttk.Button(frame1, text="PDF 파일 불러오기", bootstyle="info", command=pdfs_select_btn)
selected_files = ttk.Label(frame1, text="", bootstyle="inverse-dark")
btn_merge = ttk.Button(frame1, text="합치기", bootstyle="secondary", command=merge_pdf)
sep1 = ttk.Separator(frame1, bootstyle = "danger")

# 첫번째 row 배열
first_label.grid(row=1, column=0, sticky="nsew",padx=5, pady=5)
btn_search.grid(row=3, column=0, sticky="ew", padx=5, pady=5)
selected_files.grid(row=3, column=1, sticky="ns",padx=5, pady=5)
btn_merge.grid(row=3, column=2, sticky="ew", padx=5, pady=5 )

# 2번째 row
sep1 = ttk.Separator(frame1, bootstyle = "danger")
sec_label = Label(frame1, text = "PDF 파일을 선택해주세요.")
btn_search2 = ttk.Button(frame1, text="PPT 파일 불러오기", bootstyle="info", command=ppts_select_btn)
selected_files2 = ttk.Label(frame1, text="", bootstyle="inverse-dark")
btn_change = ttk.Button(frame1, text="PDF 변환하기", bootstyle="secondary", command=ppt_to_pdf)


# 2번째 row 배열
sep1.grid(row=4, column=0, sticky="ew",pady=5)
sec_label.grid(row=6, column=0, sticky="nsew",padx=5, pady=5)
btn_search2.grid(row=6, column=0, sticky="ew", padx=5, pady=5)
selected_files2.grid(row=6, column=1, sticky="ns",padx=5, pady=5)
btn_change.grid(row=6, column=2, sticky="ew", padx=5, pady=5 )


root.mainloop()