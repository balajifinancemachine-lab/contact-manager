@echo off
title VCF Intelligent Contact Manager Pro
echo ========================================================
echo   VCF Intelligent Contact Manager Pro - Standalone
echo ========================================================
echo.

python -m pip install -q -r requirements.txt
if %ERRORLEVEL% NEQ 0 (
    echo [WARNING] Could not install some dependencies. Attempting self-healing start anyway...
)

echo Starting application...
start pythonw contact_manager.pyw
if %ERRORLEVEL% NEQ 0 (
    echo [FALLBACK] Starting with python console...
    python contact_manager.pyw
)
