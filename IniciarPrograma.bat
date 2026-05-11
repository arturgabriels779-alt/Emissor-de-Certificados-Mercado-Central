@echo off
REM ============================================================
REM  Emissor de Certificados - Mercado Central BH
REM  Inicia o programa com instalacao automatica de Python
REM  Desenvolvido por: Artur Gabriel Oliveira da Silva
REM ============================================================
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1

cd /d "%~dp0"

title Emissor de Certificados - Mercado Central BH

REM ============================================================
REM  1. VERIFICAR SE PYTHON ESTA INSTALADO
REM ============================================================
python --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=python"
    goto DEPS_CHECK
)

py --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=py"
    goto DEPS_CHECK
)

REM ============================================================
REM  PYTHON NAO ENCONTRADO - OFERECER INSTALACAO AUTOMATICA
REM ============================================================
cls
echo.
echo  ============================================================
echo    EMISSOR DE CERTIFICADOS - MERCADO CENTRAL BH
echo  ============================================================
echo.
echo   Python nao foi encontrado no computador.
echo   O programa precisa do Python para funcionar.
echo.
echo   Deseja instalar o Python agora, automaticamente?
echo.
echo   (sera baixado cerca de 30 MB da internet e
echo    instalado apenas para o seu usuario - sem admin)
echo.
set /p RESPOSTA="   Instalar Python automaticamente? (S/N): "

if /i "!RESPOSTA!"=="S" goto INSTALL_PYTHON
if /i "!RESPOSTA!"=="SIM" goto INSTALL_PYTHON

echo.
echo   Instalacao cancelada. Baixe manualmente em:
echo   https://www.python.org/downloads/
echo.
echo   IMPORTANTE: marque "Add Python to PATH" na instalacao.
echo.
pause
exit /b 1

:INSTALL_PYTHON
echo.
echo  ------------------------------------------------------------
echo   Baixando Python 3.12 (cerca de 30 MB)...
echo  ------------------------------------------------------------
set "PY_INSTALLER=%TEMP%\python-312-installer.exe"
set "PY_URL=https://www.python.org/ftp/python/3.12.10/python-3.12.10-amd64.exe"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; " ^
  "Invoke-WebRequest -Uri '%PY_URL%' -OutFile '%PY_INSTALLER%' -UseBasicParsing; " ^
  "exit 0 } catch { Write-Host $_.Exception.Message; exit 1 }"

if errorlevel 1 (
    echo.
    echo   ERRO: Falha ao baixar o Python.
    echo   Verifique sua conexao de internet e tente novamente,
    echo   ou baixe manualmente em: https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

if not exist "%PY_INSTALLER%" (
    echo   ERRO: Arquivo baixado nao encontrado.
    pause
    exit /b 1
)

echo.
echo  ------------------------------------------------------------
echo   Instalando Python... (2-3 minutos, aguarde)
echo  ------------------------------------------------------------
echo.

REM Instalacao silenciosa apenas para o usuario atual (nao precisa admin)
"%PY_INSTALLER%" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0 Include_launcher=1 SimpleInstall=1

if errorlevel 1 (
    echo.
    echo   ERRO: Falha na instalacao do Python.
    echo   Tente instalar manualmente executando:
    echo   %PY_INSTALLER%
    echo.
    pause
    exit /b 1
)

echo   Python instalado com sucesso.
echo.

REM Atualizar o PATH da sessao atual a partir do registro do usuario
for /f "usebackq tokens=2*" %%a in (`reg query "HKCU\Environment" /v Path 2^>nul`) do set "USER_PATH=%%b"
if defined USER_PATH set "PATH=%USER_PATH%;%PATH%"

REM Limpar instalador
del /q "%PY_INSTALLER%" >nul 2>&1

REM Reverificar
python --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=python"
    goto DEPS_CHECK
)

py --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=py"
    goto DEPS_CHECK
)

echo.
echo   Python foi instalado, mas nao esta acessivel nesta sessao.
echo   POR FAVOR:
echo   1. Feche esta janela
echo   2. Abra novamente o IniciarPrograma.bat
echo.
pause
exit /b 1

:DEPS_CHECK
REM ============================================================
REM  2. VERIFICAR DEPENDENCIAS PYTHON
REM ============================================================
%PYTHON_CMD% -c "import tkinter, pandas, docxtpl, qrcode, dotenv, PIL, docx2pdf" >nul 2>&1
if not errorlevel 1 goto RUN_APP

cls
echo.
echo  ------------------------------------------------------------
echo   Instalando dependencias pela primeira vez...
echo   (pode levar 2-5 minutos)
echo  ------------------------------------------------------------
echo.

%PYTHON_CMD% -m pip install --upgrade pip --quiet
%PYTHON_CMD% -m pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo   ERRO: Falha ao instalar dependencias.
    echo   Verifique sua conexao de internet.
    echo.
    pause
    exit /b 1
)

echo.
echo   Dependencias instaladas com sucesso.
echo.

:RUN_APP
REM ============================================================
REM  3. RODAR O APLICATIVO
REM ============================================================
%PYTHON_CMD% app.py

REM Se o app fechou com erro, pausar para o usuario ver a mensagem
if errorlevel 1 (
    echo.
    echo  ------------------------------------------------------------
    echo   O programa foi encerrado com erro.
    echo   Verifique os logs na pasta "logs\" para mais detalhes.
    echo  ------------------------------------------------------------
    echo.
    pause
)

endlocal
exit /b 0
