# -*- coding: utf-8 -*-
"""
Created on Sun Jan 28 12:21:09 2018

@author: tianzhe
"""

#import docx
from docx import Document
import os
import tkinter as tk

window = tk.Tk()
window.title('请输入需要查找的字符串')
window.geometry('400x50') 

e = tk.Entry(window, width = 80)
e.pack(side = 'top')

def search(tmp_str):
    path = '.\data'
    lookup_string = tmp_str
    
    def file_name(file_dir):   
        for root, dirs, files in os.walk(file_dir):  
            return files
    
    result = '查找：' + tmp_str
    for docx_file in file_name(path):
        document = Document(path + '\\' +docx_file)
        txt = docx_file
        for paragraph in document.paragraphs:
            if lookup_string in paragraph.text:
                txt = txt + '\n' + str(paragraph.text)
        result = result + '\n' + txt
        
    file=open('result.txt','w')  
    file.write(str(result));  
    file.close()  

def input_srting():
    var = e.get()
    search(var)

b = tk.Button(window, text='确定', command = input_srting)
b.pack(side = 'bottom')

window.mainloop()

