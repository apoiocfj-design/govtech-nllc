@echo off
chcp 65001 > nul
cls
echo ==============================================================================
echo       ASSISTENTE AUTOMÁTICO DE PUBLICAÇÃO NO RENDER.COM (100%% GRATUITO)
echo ==============================================================================
echo.
echo ✅ O Render.com é 100%% gratuito e NÃO pede cartão de crédito.
echo ✅ O repositório Git local já está 100%% configurado e commitado na sua máquina!
echo ✅ Já configuramos o 'render.yaml' e o 'Dockerfile' para deploy com 1 clique.
echo.
echo ------------------------------------------------------------------------------
echo [PASSO 1 de 2] Criar o repositório vazio no seu GitHub:
echo ------------------------------------------------------------------------------
echo Deseja abrir a página para criar o repositório no GitHub agora? (S/N)
set /p opc="Digite sua opção: "
if /i "%opc%"=="S" (
    start "" "https://github.com/new?name=govtech-nllc"
)
echo.
echo No GitHub, clique no botão verde 'Create repository' (pode ser Público ou Privado).
echo Depois, copie o link HTTPS do repositório (ex: https://github.com/seu-usuario/govtech-nllc.git).
echo.
set /p repo_url="Cole a URL do seu repositório do GitHub aqui e aperte ENTER: "

if "%repo_url%"=="" (
    echo.
    echo Nenhuma URL informada. Operação cancelada.
    pause
    exit /b
)

echo.
echo Conectando repositório e enviando os arquivos...
"C:\Program Files\Git\cmd\git.exe" remote remove origin 2>nul
"C:\Program Files\Git\cmd\git.exe" remote add origin %repo_url%
"C:\Program Files\Git\cmd\git.exe" branch -M main
"C:\Program Files\Git\cmd\git.exe" push -u origin main

if %errorlevel% neq 0 (
    echo.
    echo [AVISO] Se o GitHub pediu login, autentique na janela do navegador que se abriu.
    echo Se já autenticou, tente novamente.
) else (
    echo.
    echo ==============================================================================
    echo ✅ PROJETO ENVIADO COM SUCESSO PARA O GITHUB!
    echo ==============================================================================
    echo.
    echo [PASSO 2 de 2] Ativar no Render.com:
    echo Abrindo o Render Dashboard com o nosso Blueprint automático...
    start "" "https://dashboard.render.com/select-repo?type=web"
    echo.
    echo No Render:
    echo 1. Faça login com o Google (ricardoluizrodriguesgomes@gmail.com).
    echo 2. Selecione o repositório 'govtech-nllc' que você acabou de criar.
    echo 3. O Render vai ler nosso 'render.yaml' automaticamente. Basta clicar em 'Create Web Service'!
    echo.
    echo Em 2 a 3 minutos, o seu link 'https://govtech-nllc.onrender.com' estará ONLINE!
)

pause
