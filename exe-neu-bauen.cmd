@echo off
setlocal
cd /d "%~dp0"
title BesucherTicker neu bauen

echo.
echo   BesucherTicker.exe neu bauen
echo   =============================
echo.

rem --- Laeuft die Exe noch? Dann laesst sie sich nicht ueberschreiben. ---
rem Volle Pfade, damit nicht versehentlich gleichnamige Fremdwerkzeuge
rem aus dem PATH genommen werden (z. B. das find aus Git Bash).
"%SystemRoot%\System32\tasklist.exe" /FI "IMAGENAME eq BesucherTicker.exe" 2>nul | "%SystemRoot%\System32\find.exe" /I "BesucherTicker.exe" >nul
if not errorlevel 1 (
  echo   ABBRUCH: Der Ticker laeuft gerade.
  echo   Erst beenden: Rechtsklick aufs Taskleisten-Symbol ^> Beenden.
  echo.
  pause
  exit /b 1
)

if not exist "besucher_ticker.py" (
  echo   ABBRUCH: besucher_ticker.py nicht gefunden.
  echo   Diese Datei muss im selben Ordner liegen wie das Programm.
  echo.
  pause
  exit /b 1
)

echo   Selbsttest ...
python "%~dp0besucher_ticker.py" --selbsttest
if errorlevel 1 (
  echo.
  echo   ABBRUCH: Der Selbsttest ist fehlgeschlagen.
  echo   Erst den Fehler beheben, dann neu bauen.
  echo.
  pause
  exit /b 1
)

echo.
echo   Baue ... das dauert etwa 15 Sekunden.
echo.
python -m PyInstaller --onefile --windowed --clean --noconfirm ^
  --name BesucherTicker ^
  --icon "%~dp0icon.ico" ^
  --distpath "%~dp0dist" ^
  --workpath "%~dp0build" ^
  --specpath "%~dp0build" ^
  "%~dp0besucher_ticker.py" >"%~dp0build-protokoll.txt" 2>&1

if errorlevel 1 (
  echo   FEHLER beim Bauen.
  echo   Einzelheiten stehen in build-protokoll.txt
  echo.
  pause
  exit /b 1
)

if not exist "dist\BesucherTicker.exe" (
  echo   FEHLER: Es wurde keine Exe erzeugt.
  echo   Einzelheiten stehen in build-protokoll.txt
  echo.
  pause
  exit /b 1
)

rem --- Die Exe MUSS neben der config.json liegen, sonst findet sie das Token nicht. ---
move /Y "dist\BesucherTicker.exe" "%~dp0" >nul
rmdir /S /Q "build" 2>nul
rmdir /S /Q "dist" 2>nul
del /Q "%~dp0build-protokoll.txt" 2>nul

echo   Fertig. BesucherTicker.exe ist neu gebaut.
echo.
pause
exit /b 0
