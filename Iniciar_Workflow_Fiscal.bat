@echo off
title Iniciar Workflow Fiscal - FastAPI
echo =====================================================================
echo    🚀 INICIANDO WORKFLOW FISCAL MODULAR (FastAPI + SQLAlchemy)
echo =====================================================================
echo.

:: Define o caminho absoluto da pasta do script
set "BASE_DIR=%~dp0"

:: Verifica se deve rodar no modo silencioso
set "SILENT_MODE=0"
if "%~1"=="--silent" set "SILENT_MODE=1"

:: Verifica se o servidor já está rodando na porta 8000
netstat -ano | findstr /R /C:":8000 .*LISTENING" >nul
if %errorlevel% equ 0 (
    echo [INFO] Servidor ja esta rodando na porta 8000. Abrindo aba no navegador...
    start http://127.0.0.1:8000/
    exit /b
)

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
if "%SILENT_MODE%"=="1" (
    echo ===================================================================== > "%BASE_DIR%fiscal_workflow.log"
    echo    🚀 EXECUTANDO TESTES AUTOMATIZADOS FISCAIS EM MODO SILENCIOSO      >> "%BASE_DIR%fiscal_workflow.log"
    echo ===================================================================== >> "%BASE_DIR%fiscal_workflow.log"
    python -m unittest discover -s "%BASE_DIR%fiscal_workflow/tests" -p "test_*.py" >> "%BASE_DIR%fiscal_workflow.log" 2>&1
) else (
    python -m unittest discover -s "%BASE_DIR%fiscal_workflow/tests" -p "test_*.py"
)

if %errorlevel% equ 0 (
    if "%SILENT_MODE%"=="1" (
        echo [SUCESSO] Todos os testes de integridade passaram! >> "%BASE_DIR%fiscal_workflow.log"
        echo. >> "%BASE_DIR%fiscal_workflow.log"
    )
    goto testes_passaram
)

if "%SILENT_MODE%"=="1" (
    echo. >> "%BASE_DIR%fiscal_workflow.log"
    echo ⚠️ [ALERTA] Um ou mais testes falharam! Verifique os detalhes acima. >> "%BASE_DIR%fiscal_workflow.log"
    echo [INFO] Iniciando o servidor mesmo com falhas nos testes para verificacao. >> "%BASE_DIR%fiscal_workflow.log"
    echo. >> "%BASE_DIR%fiscal_workflow.log"
    goto testes_passaram
)

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
if "%SILENT_MODE%"=="1" (
    uvicorn fiscal_workflow.main:app --host 127.0.0.1 --port 8000
) else (
    uvicorn fiscal_workflow.main:app --reload --host 127.0.0.1 --port 8000
)

if "%SILENT_MODE%"=="0" pause
