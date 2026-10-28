@echo off
setlocal
cd /d "%~dp0"
echo ================================================
echo   ORGANIZADOR DE FACTURAS V19 - CREAR EXE
echo ================================================
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo ERROR instalando dependencias.
  pause
  exit /b 1
)
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
python -m PyInstaller --noconfirm --clean --onedir --windowed --name OrganizadorFacturasV19 organizador_facturas.py
if errorlevel 1 (
  echo.
  echo ERROR creando el EXE.
  pause
  exit /b 1
)
copy /Y config.json dist\OrganizadorFacturasV19\config.json >nul
if not exist dist\OrganizadorFacturasV19\versiones_configuracion mkdir dist\OrganizadorFacturasV19\versiones_configuracion
copy /Y README_PORTABLE.txt dist\OrganizadorFacturasV19\README_PORTABLE.txt >nul

echo.
echo EXE creado en:
echo %cd%\dist\OrganizadorFacturasV19\OrganizadorFacturasV19.exe
echo.
start "" explorer "%cd%\dist\OrganizadorFacturasV19"
pause
