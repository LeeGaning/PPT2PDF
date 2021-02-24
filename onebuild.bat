@REM rmdir build /S /Q 
@REM rmdir dist /S /Q 
.\venv\Scripts\pyinstaller --version-file file_version_info.txt -F PPT2PDF.py -i pdf.ico