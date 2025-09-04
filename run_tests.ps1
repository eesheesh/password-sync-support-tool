# This script runs the Pester tests for the PasswordSyncSupportTool.
Set-Location /app
Import-Module Pester
Invoke-Pester -Path "PasswordSyncSupportTool.Tests.ps1" -Output Diagnostic
