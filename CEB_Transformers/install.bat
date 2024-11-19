@echo off
SET "PYTHON_VERSION=3.10.4"
SET "EXE_NAME=CEB_Transformers.exe"
SET "ICON_NAME=assets.ico"
SET "README_FILE=Readme.pdf"
SET "INSTALL_DIR=%~dp0"  REM Set INSTALL_DIR to the directory where the batch file is located

REM Check if the script is running as Administrator
NET SESSION >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo This script requires administrator privileges. Please grant them.
    echo Restarting script with administrator privileges...
    timeout /t 2
    REM Relaunch the batch file as administrator using PowerShell
    powershell -Command "Start-Process '%~dp0install.bat' -Verb runAs"
    exit /b
)

REM Step 1: Install Python if it's not already installed
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo Installing Python %PYTHON_VERSION%...
    REM Download and install Python (modify URL if needed)
    curl -o "%INSTALL_DIR%python_installer.exe" https://www.python.org/ftp/python/%PYTHON_VERSION%/python-%PYTHON_VERSION%-amd64.exe
    "%INSTALL_DIR%python_installer.exe" /quiet InstallAllUsers=1 PrependPath=1
    timeout /t 5
)

REM Step 2: Install required Python packages
echo Installing Python packages...
"%INSTALL_DIR%Scripts\pip.exe" install pandas openpyxl tk Pillow tkcalendar pyinstaller

REM Step 3: Create desktop shortcut for the .exe file
echo Creating desktop shortcut...
SET "DESKTOP=%USERPROFILE%\Desktop"
SET "STARTMENU=%APPDATA%\Microsoft\Windows\Start Menu\Programs"
REM Create shortcut on Desktop
echo Set oWS = WScript.CreateObject("WScript.Shell") > "%TEMP%\create_shortcut.vbs"
echo Set oLink = oWS.CreateShortcut("%DESKTOP%\CEB_Transformers.lnk") >> "%TEMP%\create_shortcut.vbs"
echo oLink.TargetPath = "%INSTALL_DIR%%EXE_NAME%" >> "%TEMP%\create_shortcut.vbs"
echo oLink.IconLocation = "%INSTALL_DIR%%ICON_NAME%" >> "%TEMP%\create_shortcut.vbs"
echo oLink.Save >> "%TEMP%\create_shortcut.vbs"
cscript //nologo "%TEMP%\create_shortcut.vbs"
del "%TEMP%\create_shortcut.vbs"

REM Create shortcut in Start Menu
echo Creating start menu shortcut...
echo Set oWS = WScript.CreateObject("WScript.Shell") > "%TEMP%\create_shortcut_startmenu.vbs"
echo Set oLink = oWS.CreateShortcut("%STARTMENU%\CEB_Transformers.lnk") >> "%TEMP%\create_shortcut_startmenu.vbs"
echo oLink.TargetPath = "%INSTALL_DIR%%EXE_NAME%" >> "%TEMP%\create_shortcut_startmenu.vbs"
echo oLink.IconLocation = "%INSTALL_DIR%%ICON_NAME%" >> "%TEMP%\create_shortcut_startmenu.vbs"
echo oLink.Save >> "%TEMP%\create_shortcut_startmenu.vbs"
cscript //nologo "%TEMP%\create_shortcut_startmenu.vbs"
del "%TEMP%\create_shortcut_startmenu.vbs"

REM Step 4: Open the README.pdf file
echo Opening README file...
start "" "%INSTALL_DIR%%README_FILE%"

REM Step 5: Start the application (first time)
echo Starting the application for the first time...
start "" "%INSTALL_DIR%%EXE_NAME%"

REM Step 6: Clean up
echo Installation complete.
pause
