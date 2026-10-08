@echo off
rem Atualiza os modelos embutidos no Gerador de Declaracoes.
rem Uso: atualizar-modelos.bat                       (embute os .docx da pasta "modelos")
rem      atualizar-modelos.bat -ConverterOriginais   (converte os .doc originais e embute)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0atualizar-modelos.ps1" %*
echo.
pause
