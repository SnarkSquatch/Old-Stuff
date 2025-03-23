@echo off
:: Change PowerShell execution policy to Unrestricted
PowerShell -Command "Set-ExecutionPolicy Bypass -Scope Process -Force"

:: Define the path to the VBScript
set "vbsFile=%~dp0ScanDrive.vbs"

:: Check if the VBScript file exists
if not exist "%vbsFile%" (
    echo VBScript file not found: %vbsFile%
    exit /b 1
)

:: Run the VBScript with elevated privileges
PowerShell -Command "Start-Process 'wscript.exe' -ArgumentList '%vbsFile%' -Verb RunAs"
