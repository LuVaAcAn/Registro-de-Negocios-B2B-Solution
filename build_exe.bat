@echo off
REM ============================================================
REM  Genera el .exe de "Registro de Negocios"
REM  Ejecutar este archivo en Windows, en la misma carpeta que
REM  app.py (doble clic o "build_exe.bat" desde una consola).
REM ============================================================

echo.
echo === Paso 1/3: Instalando dependencias (openpyxl, reportlab, pyinstaller) ===
python -m pip install --upgrade pip
python -m pip install openpyxl reportlab pyinstaller

echo.
echo === Paso 2/3: Generando el ejecutable (.exe) ===
REM --onefile      -> un solo archivo .exe portable
REM --windowed     -> sin consola negra detras de la ventana
REM --name         -> nombre del programa final
REM --clean        -> evita mezclar builds anteriores (mas rapido y sin bugs raros)
python -m PyInstaller --onefile --windowed --name "RegistroDeNegocios" --clean app.py

echo.
echo === Paso 3/3: Listo ===
echo El ejecutable quedo en la carpeta "dist\RegistroDeNegocios.exe"
echo Puedes mover ese .exe a cualquier carpeta: al abrirlo creara
echo "negocios.db" y "config.json" junto a el automaticamente.
echo.
pause
