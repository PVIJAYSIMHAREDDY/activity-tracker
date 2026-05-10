@echo off
REM Build the Windows installer for Daily Activity Tracker
REM Run this script on a Windows machine from the project root folder
REM Requirements: Python 3.8+, pip, Inno Setup 6 (https://jrsoftware.org/isdl.php)

cd /d "%~dp0.."

echo =^> Installing build dependencies...
pip install pyinstaller pywebview flask openpyxl

echo =^> Converting icon to .ico...
python -c "
from PIL import Image
import os
if os.path.exists('icon.png'):
    img = Image.open('icon.png')
    img.save('icon.ico', format='ICO', sizes=[(16,16),(32,32),(48,48),(64,64),(128,128),(256,256)])
    print('icon.ico created')
else:
    print('icon.png not found, skipping icon conversion')
" 2>nul || echo Note: PIL not installed, icon conversion skipped

echo =^> Building with PyInstaller...
pyinstaller ActivityTracker.spec --clean --noconfirm

echo =^> Checking if Inno Setup is available...
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" (
    echo =^> Building installer with Inno Setup...
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" build\windows_installer.iss
    echo.
    echo Done! Installer: build\ActivityTracker_Setup.exe
) else (
    echo.
    echo Inno Setup not found. The raw build is at: dist\ActivityTracker\
    echo Download Inno Setup from https://jrsoftware.org/isdl.php to create the installer.
)

pause
