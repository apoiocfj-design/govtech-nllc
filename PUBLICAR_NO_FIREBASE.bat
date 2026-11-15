@echo off
chcp 65001 > nul
cls
echo ==============================================================================
echo    ASSISTENTE DE PUBLICAÇÃO NO GOOGLE FIREBASE & CLOUD RUN
echo ==============================================================================
echo.
echo Conta de destino: ricardoluizrodriguesgomes@gmail.com
echo.
echo Este assistente publica o seu aplicativo na nuvem do Google (Firebase Hosting + Cloud Run)
echo para que você receba um link oficial 'https://....web.app' acessível de qualquer lugar.
echo.
echo [PASSO 1] Verificação de Ferramentas do Google:
where gcloud >nul 2>nul
if %errorlevel% neq 0 (
    echo [AVISO] O 'Google Cloud SDK' (gcloud) não foi detectado no seu PATH.
    echo.
    echo Você pode instalar em 2 minutos via terminal com o comando:
    echo   winget install Google.CloudSDK
    echo Ou baixar o instalador oficial em: https://cloud.google.com/sdk/docs/install
    echo.
) else (
    echo [OK] Google Cloud SDK (gcloud) detectado.
)

where firebase >nul 2>nul
if %errorlevel% neq 0 (
    echo [AVISO] O 'Firebase CLI' não foi detectado no seu PATH.
    echo Você pode instalar via npm: npm install -g firebase-tools
) else (
    echo [OK] Firebase CLI detectado.
)
echo.
echo ==============================================================================
echo Deseja ver o passo a passo completo de publicação com sua conta gratuita? (S/N)
set /p opc="Digite sua opção: "
if /i "%opc%"=="S" (
    start "" "notepad.exe" "%~dp0COMO_PUBLICAR_NO_FIREBASE.md"
)
pause
