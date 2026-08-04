@echo off
call .venv\Scripts\activate.bat
pyinstaller --onefile --windowed --name "Extrator Cartao Ponto" app.py
pause