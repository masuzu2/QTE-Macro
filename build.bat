@echo off
title Build QTE Macro
color 0A
echo.
echo  ========================================
echo    QTE MACRO - Build Tool
echo  ========================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python not found!
    echo  Download: https://www.python.org/downloads/
    pause
    exit /b
)

echo  [1/3] Installing dependencies...
pip install -r requirements.txt --quiet --upgrade

echo  [2/3] Creating folders...
if not exist "scripts\templates" mkdir "scripts\templates"

echo  [3/3] Building .exe ...
echo.

pyinstaller --noconfirm --onefile --windowed ^
    --name "QTE_Macro" ^
    --hidden-import=pynput ^
    --hidden-import=pynput.keyboard ^
    --hidden-import=pynput.keyboard._win32 ^
    --hidden-import=pynput.mouse ^
    --hidden-import=pynput.mouse._win32 ^
    --hidden-import=PIL ^
    --hidden-import=PIL.Image ^
    --hidden-import=PIL.ImageTk ^
    --hidden-import=PIL.ImageDraw ^
    --hidden-import=numpy ^
    --hidden-import=cv2 ^
    --hidden-import=pytesseract ^
    --hidden-import=mss ^
    --hidden-import=winsound ^
    --collect-all mss ^
    scripts/minigame_macro.py

echo.
if exist "dist\QTE_Macro.exe" (
    if not exist "dist\templates" mkdir "dist\templates"
    copy test_qte_game.html dist\ >nul
    echo  ========================================
    echo    BUILD SUCCESS!
    echo    File: dist\QTE_Macro.exe
    echo  ========================================
    explorer dist
) else (
    echo  [ERROR] Build failed!
)
echo.
pause
