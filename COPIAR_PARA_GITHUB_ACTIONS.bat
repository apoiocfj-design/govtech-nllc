@echo off
echo ================================================================
echo ATIVADOR DE CI/CD GITHUB ACTIONS (GOVTECH NLLC)
echo ================================================================
echo Certifique-se de que seu Personal Access Token no GitHub possui
echo o escopo [workflow] marcado em Settings > Developer Settings.
echo.
if not exist .github\workflows mkdir .github\workflows
copy /Y ci-cd\workflows\* .github\workflows\
echo.
echo Arquivos copiados com sucesso para .github\workflows\!
echo Agora basta rodar:
echo   git add .github/workflows
echo   git commit -m "ci: ativar pipelines no github actions"
echo   git push origin main
echo ================================================================
pause
