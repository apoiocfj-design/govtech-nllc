@echo off
chcp 65001 > nul
cls

echo Abrindo o GitHub para você criar o repositório 'govtech-nllc'...
start "" "https://github.com/new?name=govtech-nllc"

powershell -NoProfile -Command ^
  "Add-Type -AssemblyName Microsoft.VisualBasic; " ^
  "$repo = [Microsoft.VisualBasic.Interaction]::InputBox('1. No GitHub aberto, clique no botão verde Create repository.`n2. Copie a URL HTTPS do repositório gerado e cole abaixo:', 'GovTech NLLC - Publicar no GitHub e Render', ''); " ^
  "if ([string]::IsNullOrWhiteSpace($repo)) { exit 1 }; " ^
  "Set-Location 'C:\Users\CFJ-SECOM\.gemini\antigravity\scratch\govtech-nllc'; " ^
  "& 'C:\Program Files\Git\cmd\git.exe' remote remove origin 2>$null; " ^
  "& 'C:\Program Files\Git\cmd\git.exe' remote add origin $repo; " ^
  "& 'C:\Program Files\Git\cmd\git.exe' branch -M main; " ^
  "& 'C:\Program Files\Git\cmd\git.exe' push -u origin main; " ^
  "Start-Process 'https://dashboard.render.com/select-repo?type=web'"

pause
