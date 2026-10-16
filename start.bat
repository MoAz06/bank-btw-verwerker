@echo off
rem Start Kwartaal met de standaard Python 3 van de py-launcher (waar pywebview + deps op staan).
rem Dubbelklikken kan; geen zwart terminalvenster dankzij de windowed launcher (pyw).
pyw -3 "%~dp0app.py"
if errorlevel 1 (
  echo.
  echo Kon de app niet starten.
  echo Installeer eenmalig de vereisten:  py -3 -m pip install -r "%~dp0requirements.txt"
  pause
)
