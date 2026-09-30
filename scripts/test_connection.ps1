<#
.SYNOPSIS
    Tests Microsoft OneNote Desktop COM and Backup connectivity for OneNote MCP Server.

.DESCRIPTION
    Verifies that:
    1. OneNote desktop process is running and visible.
    2. OneNote COM Automation object can be instantiated.
    3. OneNote live hierarchy can be queried via COM.
    4. OneNote backup directory exists and contains .one files.
    5. PowerShell execution environment meets Single-Threaded Apartment (STA) requirements.

.EXAMPLE
    .\test_connection.ps1
#>

[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "    Microsoft OneNote MCP Environment Diagnostics Tool     " -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# 1. Check PowerShell Apartment State
$aptState = [System.Threading.Thread]::CurrentThread.GetApartmentState()
if ($aptState -eq "STA") {
    Write-Host "[OK] PowerShell Apartment State: STA (Required for OneNote COM)" -ForegroundColor Green
} else {
    Write-Host "[WARN] PowerShell Apartment State: $aptState (STA is recommended for COM stability)" -ForegroundColor Yellow
}

# 2. Check OneNote Desktop Process
$oneProcesses = Get-Process -Name "ONENOTE" -ErrorAction SilentlyContinue
if ($oneProcesses) {
    $uiProc = $oneProcesses | Where-Object { $_.MainWindowHandle -ne 0 }
    $zombieProc = $oneProcesses | Where-Object { $_.MainWindowHandle -eq 0 }

    if ($uiProc) {
        Write-Host "[OK] OneNote desktop UI process is running (PID: $($uiProc.Id))" -ForegroundColor Green
    } else {
        Write-Host "[WARN] OneNote is running but only as background headless process." -ForegroundColor Yellow
    }

    if ($zombieProc) {
        Write-Host "[INFO] Headless background instances detected: $($zombieProc.Count). These can safely be cleaned up if locks occur." -ForegroundColor Gray
    }
} else {
    Write-Host "[WARN] OneNote desktop process is NOT running. COM calls will auto-launch OneNote in background." -ForegroundColor Yellow
}

# 3. Test COM Hierarchy Call
Write-Host "`nTesting OneNote COM API connection..." -ForegroundColor Cyan
try {
    $onenote = New-Object -ComObject OneNote.Application
    $hierarchyXml = ""
    $onenote.GetHierarchy("", 3, [ref]$hierarchyXml)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($onenote) | Out-Null

    if ($hierarchyXml -and $hierarchyXml.Length -gt 50) {
        [xml]$doc = $hierarchyXml
        $notebooks = $doc.Notebooks.Notebook
        $nbCount = if ($notebooks) { $notebooks.Count } else { 0 }
        Write-Host "[OK] COM GetHierarchy succeeded! Found $nbCount active notebook(s)." -ForegroundColor Green

        foreach ($nb in $notebooks) {
            Write-Host "     - Notebook: $($nb.name)" -ForegroundColor White
            if ($nb.Section) {
                foreach ($sec in $nb.Section) {
                    Write-Host "         * Section: $($sec.name)" -ForegroundColor Gray
                }
            }
        }
    } else {
        Write-Host "[WARN] COM responded but hierarchy XML was empty." -ForegroundColor Yellow
    }
} catch {
    Write-Host "[FAIL] OneNote COM Automation failed: $_" -ForegroundColor Red
    Write-Host "       Check if OneNote has a modal dialog or password prompt open." -ForegroundColor Yellow
}

# 4. Check Backup Directory (for pyOneNote read tools)
Write-Host "`nChecking OneNote local backup files..." -ForegroundColor Cyan
$backupPath = "$env:LOCALAPPDATA\Microsoft\OneNote\16.0\Backup"
if (Test-Path $backupPath) {
    $oneFiles = Get-ChildItem -Path $backupPath -Filter "*.one" -Recurse -ErrorAction SilentlyContinue
    $fileCount = if ($oneFiles) { $oneFiles.Count } else { 0 }
    Write-Host "[OK] Backup directory found: $backupPath" -ForegroundColor Green
    Write-Host "[OK] Discovered $fileCount backup (.one) file(s) for offline reading." -ForegroundColor Green
} else {
    Write-Host "[WARN] Backup path not found at $backupPath" -ForegroundColor Yellow
    Write-Host "       In OneNote: Go to File -> Options -> Save & Backup -> Back Up All Notebooks Now" -ForegroundColor Gray
}

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "Diagnostics Complete!" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
