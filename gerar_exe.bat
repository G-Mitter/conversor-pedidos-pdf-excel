@echo off
title Gerar Conversor_de_Pedidos.exe
cd /d "%~dp0"
echo ============================================
echo   Gerando o executavel do Conversor...
echo ============================================
echo.
where python >nul 2>nul || (echo ERRO: Python nao encontrado. Instale em https://www.python.org/downloads/ & goto fim)
echo [1/2] Instalando/atualizando dependencias (pode levar alguns minutos)...
python -m pip install --upgrade pymupdf pdfplumber pypdf openpyxl pillow numpy rapidocr-onnxruntime pyinstaller || (echo. & echo ERRO ao instalar dependencias. Veja a mensagem acima. & goto fim)
echo.
echo [2/2] Gerando o .exe com PyInstaller (leva 1 a 3 minutos)...
python -m PyInstaller Conversor_de_Pedidos.spec --noconfirm || (echo. & echo ERRO ao gerar o executavel. Veja a mensagem acima. & goto fim)
echo.
echo Pronto! Executavel gerado em:
echo    %~dp0dist\Conversor_de_Pedidos.exe
:fim
echo.
pause
