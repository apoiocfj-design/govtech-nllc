@echo off
chcp 65001 > nul
cls
echo ============================================================================
echo   INICIANDO APLICATIVO UNIFICADO GOVTECH NLLC (LEI 14.133/2021 E IN 65/2021)
echo ============================================================================
echo.
echo 👉 Seu link local:                  http://localhost:8080
echo 👉 Link para colegas na mesma rede: http://192.168.1.5:8080
echo.
echo Abrindo seu navegador automaticamente...
start "" "http://localhost:8080"
"C:\Users\CFJ-SECOM\.gemini\antigravity\scratch\pesquisa-precos\.venv\Scripts\python.exe" "C:\Users\CFJ-SECOM\.gemini\antigravity\scratch\govtech-nllc\backend\main.py"
pause
