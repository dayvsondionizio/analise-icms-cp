@echo off
TITLE Iniciando Analise ICMS CP...

echo Verificando dependencias...
IF NOT EXIST "node_modules" (
    echo Instalando dependencias pela primeira vez (isso pode demorar alguns minutos)...
    call npm install
)

echo Iniciando o servidor...
start http://localhost:3000
call npm run dev

pause
