@echo off

cd "%~dp0"
IF NOT EXIST "data\3.cer" (
    echo .
    echo ������Ҫ��װ��֤���ļ�ȱʧ��
    echo ����������ѹ�������ļ�����ʹ�ñ����ߡ�
    echo .
    echo �밴�ո������������Ƴ�...
    pause >nul
    exit /b
)

:: BatchGotAdmin
:-------------------------------------
REM  --> Check for permissions
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"

REM --> If error flag set, we do not have admin.
if '%errorlevel%' NEQ '0' (
    echo.
    echo ��Ҫ����ԱȨ�޲��ܰ�װ��֤��
    echo.
    echo �밴�ո������������Ի�ȡ����ԱȨ��...
    pause >nul
    goto UACPrompt
) else ( goto gotAdmin )

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    exit /B

:gotAdmin
    if exist "%temp%\getadmin.vbs" ( del "%temp%\getadmin.vbs" )
    pushd "%CD%"
    CD /D "%~dp0"
:--------------------------------------

echo.
certutil -addstore -f "ROOT" "data\3.cer"
echo.
echo �밴�ո������������Ƴ�...
pause >nul
