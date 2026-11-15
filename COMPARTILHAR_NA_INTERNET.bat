@echo off
chcp 65001 >nul
cls
echo ==============================================================================
echo       SISTEMA GOVTECH NLLC • COMPARTILHAMENTO DO LINK PARA QUALQUER PESSOA
echo ==============================================================================
echo.
echo 1. ACESSO NA REDE LOCAL / WI-FI (Sem precisar de internet):
echo    Qualquer colega na mesma rede Wi-Fi ou cabo de rede pode acessar:
echo    👉 http://192.168.1.5:8080
echo.
echo 2. LINK PUBLICO PELA INTERNET (Para qualquer pessoa no Brasil - Celular ou PC):
echo    Voce pode usar o Cloudflare Tunnel oficial (gratuito e sem cadastro).
echo    Deseja iniciar um Link Publico Seguro agora? (S/N)
set /p opc="Digite sua opcao: "
if /i "%opc%"=="S" (
    echo.
    echo Iniciando túnel seguro da Cloudflare para a porta 8080...
    echo Copie o link 'https://....trycloudflare.com' que aparecera abaixo e envie aos colegas.
    echo.
    npx -y cloudflared tunnel --url http://localhost:8080
) else (
    echo.
    echo Para acessar localmente na sua propria maquina:
    echo 👉 http://localhost:8080
    pause
)
