@echo off
setlocal enabledelayedexpansion

:: Change directory to the project root (parent directory of 'scripts')
cd /d "%~dp0.."

echo ========================================
echo   DanteToExcel Release Builder
echo ========================================
echo.

:: Clean and create dist directory
if exist dist (
    echo Cleaning old dist directory...
    rmdir /s /q dist
)
mkdir dist

:: Save current environment variables
set ORIGINAL_GOOS=%GOOS%
set ORIGINAL_GOARCH=%GOARCH%

echo.
echo 1. Building Windows (x64)...
set GOOS=windows
set GOARCH=amd64
go build -ldflags="-s -w" -o dist/DanteToExcel_windows_x64.exe src/main.go
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Failed to build Windows binary.
    goto error
)

echo.
echo 2. Building macOS (Intel)...
set GOOS=darwin
set GOARCH=amd64
go build -ldflags="-s -w" -o dist/DanteToExcel_macOS_Intel src/main.go
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Failed to build macOS Intel binary.
    goto error
)

echo.
echo 3. Building macOS (Apple Silicon)...
set GOOS=darwin
set GOARCH=arm64
go build -ldflags="-s -w" -o dist/DanteToExcel_macOS_AppleSilicon src/main.go
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Failed to build macOS Apple Silicon binary.
    goto error
)

:: Restore environment variables
set GOOS=%ORIGINAL_GOOS%
set GOARCH=%ORIGINAL_GOARCH%

echo.
echo 4. Packaging releases into Zip archives...

:: Create Windows zip
powershell -NoProfile -Command "Compress-Archive -Path 'dist/DanteToExcel_windows_x64.exe', 'docs/manual_JP.md', 'docs/manual_EN.md' -DestinationPath 'dist/DanteToExcel_windows_x64.zip' -Force"
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Failed to package Windows release.
    goto error
)

:: Create macOS Intel zip
powershell -NoProfile -Command "Compress-Archive -Path 'dist/DanteToExcel_macOS_Intel', 'docs/manual_JP.md', 'docs/manual_EN.md' -DestinationPath 'dist/DanteToExcel_macOS_Intel.zip' -Force"
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Failed to package macOS Intel release.
    goto error
)

:: Create macOS Apple Silicon zip
powershell -NoProfile -Command "Compress-Archive -Path 'dist/DanteToExcel_macOS_AppleSilicon', 'docs/manual_JP.md', 'docs/manual_EN.md' -DestinationPath 'dist/DanteToExcel_macOS_AppleSilicon.zip' -Force"
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Failed to package macOS Apple Silicon release.
    goto error
)

echo.
echo ========================================
echo   Build completed successfully.
echo   Check output in the 'dist' folder.
echo ========================================
endlocal
exit /b 0

:error
:: Restore environment variables
set GOOS=%ORIGINAL_GOOS%
set GOARCH=%ORIGINAL_GOARCH%
echo [ERROR] Build failed.
endlocal
exit /b 1
