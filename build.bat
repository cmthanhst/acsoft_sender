@echo off
echo =================================================================
echo            DONG DEP TRUOC KHI BUILD
echo =================================================================
echo.
REM SỬA: Đảm bảo file .exe cũ không còn chạy
taskkill /f /im VietTinAutoSender.exe >nul 2>nul
echo [INFO] Da gui lenh dong ung dung (neu dang chay).

REM SỬA: Thêm độ trễ ngắn để đảm bảo OS giải phóng file lock
echo [INFO] Dang cho 2 giay de he dieu hanh giai phong file...
timeout /t 2 /nobreak >nul

IF EXIST "dist" ( rmdir /s /q dist )
IF EXIST "build" ( rmdir /s /q build )
IF EXIST "VietTinAutoSender.spec" ( del VietTinAutoSender.spec )
echo [INFO] Da xoa cac thu muc build va dist cu (neu co).

REM SỬA: Kiểm tra lại xem thư mục dist đã được xóa thành công chưa
IF EXIST "dist" (
    echo.
    echo [LOI] Khong the xoa thu muc 'dist'. Co the file .exe van dang bi khoa.
    echo Vui long thu dong ung dung bang tay qua Task Manager hoac tam tat phan mem diet virus roi chay lai.
    pause
    exit /b
)
echo.

echo =================================================================
echo            KICH BAN DONG GOI UNG DUNG .EXE
echo =================================================================
echo.

REM Kiem tra xem Python co trong PATH khong
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [LOI] Khong tim thay Python trong PATH. Vui long cai dat Python va them vao PATH.
    pause
    exit /b
)

REM SỬA: Sử dụng 'python -m pip' để tránh lỗi PATH
echo [INFO] Dang kiem tra va cai dat PyInstaller (su dung 'python -m pip')...
REM SỬA: Cài đặt thêm pyinstaller-hooks-contrib để hỗ trợ các thư viện phức tạp
python -m pip install pyinstaller pyinstaller-hooks-contrib

echo.
echo [INFO] Bat dau qua trinh dong goi. Viec nay co the mat vai phut...
echo [INFO] Cac file tai nguyen se duoc them vao:
echo        - dark.qss
echo        - light.qss
echo        - logo_base64.txt
echo.

REM SỬA: Sử dụng 'python -m PyInstaller' để tránh lỗi PATH
REM SỬA: Thêm các hidden-import cho tất cả các thư viện có khả năng bị bỏ sót
python -m PyInstaller --name "VietTinAutoSender" --onefile --windowed --icon="logo.ico" ^
--add-data "dark.qss;." ^
--add-data "light.qss;." ^
--add-data "logo_base64.txt;." ^
--hidden-import=PySide6 ^
--hidden-import=pynput --hidden-import=pandas --hidden-import=win32gui main.py

echo.
echo =================================================================
echo [SUCCESS] Dong goi hoan tat!
echo Tim file .exe trong thu muc 'dist'.
echo =================================================================
pause