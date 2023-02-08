import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import *
from tkinter import filedialog

# 파일 불러오기

# 파일 선택
def file_select_btn():
    global root
    root.filename = filedialog.askopenfilename(initialdir = "C:/flyordig/et_work/docs",
                                                title = "PDF 파일을 골라주세요",
                                                filetypes = (("PDF","*.pdf"),("all files","*.*")))
    ssfile = root.filename
    root.withdraw()

# GUI 부분
root = ttk.Window()
root.title("pdf 합치기 도구")
frame1 = ttk.Frame(root, width=700, height=500, bootstyle = "default")
frame1.pack(fill=tk.X, expand=True)
# frame2 = ttk.Frame(root, width=700, height=100)
# frame2.pack(fill=tk.X, expand=True)

# 프레임 1의 첫번째 열
first_label = Label(frame1, text = "PDF 파일을 선택해주세요.")
btn_search = ttk.Button(frame1, text="PDF 파일 불러오기",
                        bootstyle="info", command=file_select_btn)
selected_file_list = Listbox(frame1)
btn_merge = ttk.Button(frame1, text="합치기", bootstyle="secondary")

# grid 정렬
first_label.grid(row=1, column=0, sticky="nsew",padx=5, pady=5)
btn_search.grid(row=3, column=0, sticky="ew", padx=5, pady=5)
selected_file_list.grid(row=3, column=1, sticky="ns",padx=5, pady=5)
btn_merge.grid(row=3, column=2, sticky="ew", padx=5, pady=5 )


root.mainloop()