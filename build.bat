@echo off
echo ============================================
echo  Access DB Documenter - Build Script
echo ============================================
echo.

echo [1/2] Installing / upgrading dependencies...
pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: pip install failed.
    pause & exit /b 1
)

echo.
echo [2/2] Building standalone executable...
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "AccessDBDocumenter" ^
    documenter.py

if errorlevel 1 (
    echo ERROR: PyInstaller build failed.
    pause & exit /b 1
)

echo.
echo ============================================
echo  Build complete!
echo  Output: dist\AccessDBDocumenter.exe
echo  Distribute that single .exe to students.
echo ============================================
pause
