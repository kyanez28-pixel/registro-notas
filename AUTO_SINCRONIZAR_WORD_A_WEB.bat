@echo off
title SINCRONIZADOR AUTOMATICO WORD -> WEB
color 0A
cd /d "%~dp0"
echo ======================================================================
echo    INICIANDO SINCRONIZADOR AUTOMATICO DE PLANIFICACION MICROCURRICULAR
echo    Cada vez que guardes el archivo Word (Ctrl + S), se actualizara
echo    automaticamente la web local y en GitHub Pages.
echo ======================================================================
echo.
python watch_planificacion.py
pause
