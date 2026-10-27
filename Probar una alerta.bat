@echo off
REM One-shot test launcher. Processes a SINGLE alert so the PDF layout and the
REM resulting draft can be checked before the whole backlog is consumed.
REM
REM This is NOT the daily launcher. Use "Procesar alertas.bat" for that: --limit
REM is a per-run cap, not a pace, so this file would need one double-click per
REM alert.
REM
REM Output is deliberately plain ASCII. The console codepage mangles accents, and
REM a technician reading garbled text assumes something broke.

title Clariot - Probar una alerta
cd /d "%~dp0"

echo ==================================================
echo   CLARIOT - PRUEBA: una sola alerta
echo ==================================================
echo.
echo Outlook tiene que estar abierto.
echo.
echo Esto procesa UN correo, no todos. Sirve para
echo revisar que el borrador salga bien antes de
echo soltar el resto.
echo.
echo --------------------------------------------------

".venv\Scripts\python.exe" -m clariot --limit 1
set CODIGO=%ERRORLEVEL%

echo --------------------------------------------------
echo.
if "%CODIGO%"=="0" (
    echo   LISTO. Abre el borrador en Outlook:
    echo.
    echo     Borradores / Urgencias   o   Por enviar
    echo.
    echo   Revisa que:
    echo     - el PDF original este adjunto
    echo     - el informe en espanol se lea bien
    echo     - el destinatario: cargado si el cliente esta
    echo       en clients.yaml, vacio si no
    echo     - aparezcan [NOMBRE] y [COMPLETAR]
    echo     - al final este la firma de Outlook
    echo.
    echo   Si esta todo bien, usa "Procesar alertas"
    echo   para el resto.
) else (
    echo   TERMINO CON PROBLEMAS.
    echo   Mira el detalle arriba, o abre logs\clariot.log
)
echo.
echo ==================================================
echo.
pause
