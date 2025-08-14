@echo off
REM Script tu dong build file .exe cho ung dung Barcode Reader

echo =======================================================
echo  Bat dau qua trinh build file .exe
echo  Vui long khong tat cua so nay...
echo =======================================================
echo.

REM Buoc 1: Kich hoat moi truong ao
echo [1/3] Kich hoat moi truong ao (.venv)...
call .\.venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo Loi: Khong the kich hoat moi truong ao. Hay kiem tra lai duong dan.
    pause
    exit /b
)
echo Moi truong ao da duoc kich hoat.
echo.

REM Buoc 2: Kiem tra va cai dat PyInstaller neu can
echo [2/3] Kiem tra PyInstaller...
pip show pyinstaller > nul 2>&1
if %errorlevel% neq 0 (
    echo PyInstaller chua duoc cai dat. Dang tu dong cai dat...
    pip install pyinstaller
) else (
    echo PyInstaller da duoc cai dat.
)
echo.


REM Buoc 3: Chay PyInstaller voi cac tuy chon can thiet
echo [3/3] Dang build file .exe... Qua trinh nay co the mat vai phut.
pyinstaller --name "BarcodeScanner" ^
            --onefile ^
            --windowed ^
            --icon="Logo.ico" ^
            --add-binary ".venv\Lib\site-packages\pyzbar\libzbar-64.dll;pyzbar" ^
            barcode_reader.py

echo.
echo =======================================================
echo  QUA TRINH BUILD HOAN TAT!
echo.
echo  File "BarcodeScanner.exe" da duoc tao trong thu muc 'dist'.
echo =======================================================
echo.
pause
