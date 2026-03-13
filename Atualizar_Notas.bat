@echo off
title Importador de Notas Fiscais - Cabral e Sousa
echo ========================================================
echo      ATUALIZANDO BANCO DE DADOS DE NOTAS FISCAIS
echo ========================================================
echo.
echo Por favor, aguarde enquanto o sistema verifica notas novas...
echo.

"C:\stenio\Programs\Conversor de XML\Conversor_de_Xml\venv\Scripts\python.exe" "C:\stenio\Programs\Conversor de XML\Conversor_de_Xml\importador_nfe.py"

echo.
echo ========================================================
echo Processo finalizado! Voce ja pode fechar esta janela ou abrir o Excel.
pause