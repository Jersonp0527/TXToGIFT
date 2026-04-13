@echo off
cd /d "%~dp0"
python -m pip install --upgrade pip
python -m pip install -e .
if errorlevel 1 exit /b 1
echo.
echo Instalado. Ahora puedes ejecutar: txt-to-gift archivo.docx [salida.txt]
