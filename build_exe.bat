@echo off
REM ============================================================
REM  Genera el .exe de "Registro de Negocios"
REM  Ejecutar este archivo en Windows, en la misma carpeta que
REM  app.py (doble clic o "build_exe.bat" desde una consola).
REM ============================================================

echo.
echo === Paso 1/3: Instalando dependencias ===
echo (openpyxl, reportlab, cryptography, matplotlib, pyinstaller)
python -m pip install --upgrade pip
python -m pip install openpyxl reportlab cryptography matplotlib pyinstaller

echo.
echo === Paso 2/3: Generando el ejecutable (.exe) ===
REM --onefile      -> un solo archivo .exe portable
REM --windowed     -> sin consola negra detras de la ventana
REM --name         -> nombre del programa final
REM --clean        -> evita mezclar builds anteriores (mas rapido y sin bugs raros)
REM Nota: con matplotlib el .exe pesa bastante mas (~150-250 MB) y el
REM primer arranque tarda un poco mas. Es normal.
python -m PyInstaller --onefile --windowed --name "RegistroDeNegocios" --clean app.py

echo.
echo === Paso 3/3: Listo ===
echo El ejecutable quedo en la carpeta "dist\RegistroDeNegocios.exe"
echo Puedes mover ese .exe a cualquier carpeta: al abrirlo creara
echo "negocios.db.enc", "negocios.key", "config.json", "backups\" y
echo "adjuntos\" junto a el automaticamente.
echo.
echo IMPORTANTE: no borres "negocios.key" — es la clave que descifra
echo tu base de datos. Sin ella, negocios.db.enc no se puede recuperar.
echo.
pause
