# pdf 이미지 변환 부분
from pdf2image import convert_from_path
import os, sys

# 모듈 import
import re
import pyautogui as pg
# gui 부분
import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import *
from tkinter import filedialog

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)
    
# PDF to Image(JPEG) function
def convert_pdf_to_image(input_path, save_path):
    abs_input_path = os.path.abspath(input_path)
    abs_save_path = os.path.abspath(save_path)
    splits_path = abs_input_path.split('\\')
    file_name = splits_path[len(splits_path) - 1]
                      
    # split 저장할 파일명
    sub_index = len(file_name) - 4
    save_file_name = file_name[:sub_index]
    # convert pdf to images
    path = resource_path('lib\\poppler-23.01.0\\Library\\bin')
    images = convert_from_path(abs_input_path, dpi=600, thread_count=4, poppler_path = path)
      
    # 저장할 폴더 없는 경우 생성
    if (os.path.exists(abs_save_path) == False):
        os.mkdir(abs_save_path)
     
    image_count = len(images)
    for i in range(image_count):
      # Save pages as images in the pdf
      images[i].save(abs_save_path + "\\" + save_file_name + '_' + str(i + 1) +'.jpg', 'JPEG')
      

if __name__=="__main__":
    convert_pdf_to_image()