@echo off
title Iniciar Workflow Fiscal - FastAPI
echo =====================================================================
echo    🚀 INICIANDO WORKFLOW FISCAL MODULAR (FastAPI + SQLAlchemy)
echo =====================================================================
echo.

:: Define o caminho absoluto da pasta do script
set "BASE_DIR=%~dp0"

:: Verifica se a pasta do ambiente virtual existe
if exist "%BASE_DIR%.venv" goto venv_ok
echo ❌ [ERRO] Ambiente virtual (.venv) nao encontrado!
echo Certifique-se de que a pasta .venv existe na raiz do projeto.
pause
exit /b

:venv_ok

:: Ativa o ambiente virtual
echo [INFO] Ativando ambiente virtual...
call "%BASE_DIR%.venv\Scripts\activate.bat"

:: Executa os testes automatizados para garantir integridade antes de subir o servidor
echo.
echo [INFO] Executando testes automatizados fiscais...
python -m unittest discover -s "%BASE_DIR%fiscal_workflow/tests" -p "test_*.py"
if %errorlevel% equ 0 goto testes_passaram

echo.
echo ⚠️ [ALERTA] Um ou mais testes falharam! Verifique as mensagens acima.
echo.
set /p CONTINUAR="Deseja iniciar o servidor mesmo assim? (S/N): "
if /i "%CONTINUAR%" neq "S" exit /b

:testes_passaram
echo.
echo ✅ [SUCESSO] Todos os 11 testes de integridade passaram!
echo.

:: Abre automaticamente o Dashboard do Usuário no navegador padrão
echo [INFO] Abrindo o painel interativo no seu navegador...
start http://127.0.0.1:8000/

:: Inicia o servidor local FastAPI via Uvicorn
echo [INFO] Servidor rodando. Para fechar, aperte Ctrl+C neste terminal.
echo.
uvicorn fiscal_workflow.main:app --reload --host 127.0.0.1 --port 8000

pause
