@echo off
echo ===================================
echo Google Maps Scraper - Selenium Fix
echo ===================================
echo.

echo 1. Checking Python installation...
python --version
if %errorlevel% neq 0 (
    echo Error: Python is not installed or not in PATH
    echo Please install Python from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation
    pause
    exit /b 1
)

echo.
echo 2. Checking Chrome installation...
where chrome 2>nul
if %errorlevel% neq 0 (
    echo Chrome not found in PATH, checking common locations...
    
    if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" (
        echo Found Chrome in Program Files
    ) else if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" (
        echo Found Chrome in Program Files (x86)
    ) else if exist "%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe" (
        echo Found Chrome in AppData
    ) else (
        echo Error: Google Chrome is not installed or not found
        echo Please install Chrome from https://www.google.com/chrome/
        pause
        exit /b 1
    )
)

echo.
echo 3. Removing existing webdriver manager cache...
if exist "%USERPROFILE%\.wdm" (
    rmdir /s /q "%USERPROFILE%\.wdm"
    echo Removed .wdm cache directory
)

if exist "%USERPROFILE%\.cache\selenium" (
    rmdir /s /q "%USERPROFILE%\.cache\selenium"
    echo Removed selenium cache directory
)

echo.
echo 4. Reinstalling Selenium and webdriver-manager...
pip uninstall -y selenium webdriver-manager
pip install selenium==4.16.0 webdriver-manager==4.0.1

echo.
echo 5. Running the setup tool to download the correct ChromeDriver...
python install_chrome.py

echo.
echo 6. Setup complete!
echo You can now try running the script again.
echo.

pause 