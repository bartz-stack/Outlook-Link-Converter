@echo off
echo ========================================
echo Outlook Link Converter - Build Script
echo (Right-Click Context Menu Version)
echo ========================================
echo.

:: Check for command line argument
if "%1"=="clean" goto :cleanup
if "%1"=="CLEAN" goto :cleanup

:build
REM Check if virtual environment exists
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
    echo.
)

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat
echo.

REM Install/upgrade dependencies
echo Installing dependencies...
pip install --upgrade pip
pip install pyinstaller
pip install pillow
echo.

REM Clean previous builds
echo Cleaning previous builds...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "*.spec" del /q *.spec
echo.

REM === Stage 1: Build the converter exe ===
echo [1/3] Building converter executable...
if exist "Outlook365_icon.ico" (
    pyinstaller --onefile --windowed --icon=Outlook365_icon.ico --name "OutlookLinkConverter" --hidden-import=PIL --hidden-import=PIL._tkinter_finder outlook_link_converter.py
) else (
    pyinstaller --onefile --windowed --name "OutlookLinkConverter" --hidden-import=PIL --hidden-import=PIL._tkinter_finder outlook_link_converter.py
)
echo.

REM === Stage 2: Build the uninstaller exe ===
echo [2/3] Building uninstaller executable...
pyinstaller --onefile --windowed --name "Uninstall" --icon=Outlook365_icon.ico uninstaller.py
echo.

REM Copy the uninstaller to dist so the installer can find it
if exist "dist\Uninstall.exe" (
    echo Uninstaller built successfully.
) else (
    echo WARNING: Uninstaller build failed!
)

REM Copy support files to dist folder
echo Copying support files to dist...
copy gui_config.json dist\
if exist "Outlook365_icon.ico" copy Outlook365_icon.ico dist\
echo.

REM === Stage 3: Build the installer exe (bundles everything) ===
echo [3/3] Building installer executable...
pyinstaller --onefile --windowed --name "OutlookLinkConverterInstaller" ^
    --icon=Outlook365_icon.ico ^
    --hidden-import=PIL --hidden-import=PIL._tkinter_finder ^
    --add-data "dist\OutlookLinkConverter.exe;." ^
    --add-data "dist\Uninstall.exe;." ^
    --add-data "gui_config.json;." ^
    --add-data "Outlook365_icon.ico;." ^
    installer.py
echo.

echo ========================================
echo Build Complete!
echo ========================================
echo.
echo Output files in dist\:
dir /b dist
echo.
echo The file to put on the network share:
echo   dist\OutlookLinkConverterInstaller.exe
echo.
echo Users double-click it, done. No admin needed.
echo.
pause
goto :eof

:cleanup
echo ========================================
echo Cleaning Up Build Artifacts
echo ========================================
echo.

echo Deleting venv folder...
if exist "venv" (
    rmdir /s /q venv
    echo   Done - Deleted venv/
) else (
    echo   - venv/ not found
)

echo Deleting build folder...
if exist "build" (
    rmdir /s /q build
    echo   Done - Deleted build/
) else (
    echo   - build/ not found
)

echo Deleting dist folder...
if exist "dist" (
    rmdir /s /q dist
    echo   Done - Deleted dist/
) else (
    echo   - dist/ not found
)

echo Deleting __pycache__...
if exist "__pycache__" (
    rmdir /s /q __pycache__
    echo   Done - Deleted __pycache__/
) else (
    echo   - __pycache__/ not found
)

echo Deleting spec files...
if exist "*.spec" (
    del /q *.spec
    echo   Done - Deleted *.spec files
) else (
    echo   - No spec files found
)

echo.
echo ========================================
echo Cleanup Complete!
echo ========================================
echo.
pause
goto :eof