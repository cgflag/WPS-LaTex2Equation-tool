@echo off
cd /d "%~dp0.."
echo Building WPS-LaTeX2Equation.exe ...
pip install -q -e ".[gui]" pyinstaller
pyinstaller --noconfirm scripts\WPS-LaTeX2Equation.spec
if %ERRORLEVEL% equ 0 (
    echo.
    echo OK: dist\WPS-LaTeX2Equation.exe
) else (
    echo Build failed.
)
pause
