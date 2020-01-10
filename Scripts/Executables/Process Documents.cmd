@echo off
:: The following variable should be set to where your PowerShell Executable resides.
:: DO NOT USE ANY SPACES BEFORE/AFTER THE EQUALS SIGN IN THE FOLLOWING LINE.

SET scriptLocation=C:\Axis
cd %scriptLocation%
powershell.exe -executionpolicy bypass "%scriptLocation%\Process-IEP504Documents.ps1"
exit
