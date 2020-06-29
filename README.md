# docx-search-manipulation
Search and Manipulate .DOCX files with python

## 緣起
這學期 (108-2) 修大學部的 python 課程，在期末用 [`python-docx`](https://python-docx.readthedocs.io/en/latest/) 搭配 python 內建函式庫做了一個迷你專案，方便搜尋及操作 .docx 檔。

## 使用

1. 輸入原資料夾所在的路徑。例如：`C:\Users\user\Desktop\final_project_test\doc_files`。

2. 輸入關鍵字（可接受 regular expression），就可以找出含有此關鍵字的所有 .docx 檔案。

3. 輸入 `move` 或 `copy`，再輸入新路徑（欲移動或複製至哪一個資料夾），例如：`C:\Users\user\Desktop\final_project_test\new_doc_files`，就可以將含有關鍵字的 .docx 檔案移動或複製到新的資料夾。或是直接輸入 `delete`，將含有關鍵字的 .docx 檔案刪除。
