@echo off
cd /d "%~dp0"

echo =============================== >> log_buscador.txt
echo Inicio: %date% %time% >> log_buscador.txt
py -3.14 "buscador_trabajos_v2.py" >> log_buscador.txt 2>&1
echo Fin: %date% %time% >> log_buscador.txt
echo =============================== >> log_buscador.txt