@echo off
setlocal

:: --- ЭТАП 1: АВТО-ПОИСК PYTHON ---
set "PYTHON_EXE="

:: 1. Проверяем, работает ли команда python глобально (и не заглушка ли это)
python --version >nul 2>&1
if %errorlevel% EQU 0 (
    set "PYTHON_EXE=python"
    goto :Found
)

:: 2. Проверяем команду 'py' (Python Launcher)
py --version >nul 2>&1
if %errorlevel% EQU 0 (
    set "PYTHON_EXE=py"
    goto :Found
)

:: 3. Ищем в папке пользователя (используем переменную %LOCALAPPDATA%)
:: Это сработает для ЛЮБОГО пользователя, так как %LOCALAPPDATA% у всех свой.
:: Проверяем версию 3.13 (твоя) и популярные 3.12, 3.11
for %%v in (313 312 311 310) do (
    if exist "%LOCALAPPDATA%\Programs\Python\Python%%v\python.exe" (
        set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python%%v\python.exe"
        goto :Found
    )
)

:: 4. Ищем в глобальной папке Program Files
if exist "C:\Program Files\Python313\python.exe" (
    set "PYTHON_EXE=C:\Program Files\Python313\python.exe"
    goto :Found
)

:: Если ничего не нашли
echo [ERROR] Python ne nayden avtomaticheski.
echo Pozhaluysta, ubedites, chto Python ustanovlen.
pause
exit /b

:Found
:: --- ЭТАП 2: ПРАВА АДМИНИСТРАТОРА ---
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    goto UACPrompt
) else ( goto RunScript )

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    echo UAC.ShellExecute "cmd.exe", "/c ""%~s0""", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    exit /B

:RunScript
    if exist "%temp%\getadmin.vbs" ( del "%temp%\getadmin.vbs" )
    pushd "%CD%"
    CD /D "%~dp0"

    :: --- ЭТАП 3: ЗАПУСК ---
    "%PYTHON_EXE%" vizfix.py

    if %errorlevel% NEQ 0 pause
