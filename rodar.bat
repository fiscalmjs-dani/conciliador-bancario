@echo off
title Conciliador Bancario
cd /d C:\Conciliador
call venv\Scripts\activate
echo.
echo Iniciando Conciliador Bancario...
echo Aguarde o navegador abrir automaticamente.
echo.
echo Para fechar o sistema, feche esta janela.
echo.
streamlit run app.py
pause
