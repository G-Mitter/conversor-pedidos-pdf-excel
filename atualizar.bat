@echo off
title Atualizar Conversor de Pedidos
cd /d "%~dp0"
echo ============================================
echo   Atualizando a base local pelo GitHub...
echo ============================================
echo.
where git >nul 2>nul || (echo ERRO: Git nao instalado. Baixe em https://git-scm.com/download/win & goto fim)
echo [1/2] Baixando correcoes do GitHub (branch main)...
git pull origin main || (echo. & echo ERRO: nao foi possivel atualizar. Veja a mensagem acima. & goto fim)
echo.
echo [2/2] Ultima correcao baixada:
git log -1 --format="   %%h  %%ad  %%s" --date=format:"%%d/%%m/%%Y %%H:%%M"
echo.
echo Pronto! Base local atualizada.
:fim
echo.
pause
