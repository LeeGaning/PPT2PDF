rmdir build /S /Q 
rmdir dist /S /Q 
 .\venv\Scripts\pyinstaller --version-file file_version_info.txt -F PPT转换PDF.py -i pdf.ico