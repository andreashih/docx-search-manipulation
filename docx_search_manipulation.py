import os
import docx
import os.path
from os import path
from docx import Document
import re
import shutil
from shutil import copyfile

# User provides file path
path1 = '123'
while path.exists(path1) == False:
    path1 = input('Please enter the location of the original files:')  
    if path.exists(path1):
        os.chdir(path1) # Set working directory
    else:
        print('Warning: Invalid path')

# A function that extracts files that match the pattern
def search_file(filename, pattern):
    doc = docx.Document(filename)
    text = []
    for para in doc.paragraphs:
        text.append(para.text)
    if re.findall(pattern, str(text)) != []:
        return filename
    else:
        return 0

# Search docx files with content matching particular pattern
pattern = input('Please enter the keyword')
matched_files = []
for file in os.listdir():
    if not file.endswith('.docx'): continue
    x = search_file(file, pattern)
    if x != 0:
        matched_files.append(x)

# Extract matched files path
matched_files_paths = []
for matched_file in matched_files:
    path = f"{path1}\\{matched_file}"
    matched_files_paths.append(path)

# A function that manipulates files (move, copy, or delete)
def file_manipulate(mode, matched_file_path, to):
    if mode == 'move':
        shutil.move(matched_file_path, to)
    elif mode == 'copy':
        shutil.copy(matched_file_path, to)
    elif mode == 'delete':
        os.remove(matched_file_path)        
    else:
        return 'no such mode'

# User provides mode and new file path
mode = '123'
while mode not in ['move','copy', 'delete']:
    mode = input('move, copy, or delete?')
    if mode == 'delete':
        to = ''
        for matched_file_path in matched_files_paths:
                file_manipulate(mode, matched_file_path, to)  
    elif mode in ['move', 'copy']:
        to = input("Please provide the new file path")
        if path.exists(to) == False:
            print('Warning: Invalid path')
        else:
            for matched_file_path in matched_files_paths:
                file_manipulate(mode, matched_file_path, to)
    else:
        print('Warning: Invalid mode')

