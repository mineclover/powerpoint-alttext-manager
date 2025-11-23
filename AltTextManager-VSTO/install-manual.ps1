# PowerPoint AltTextManager - Manual Installation Script
# This script installs the VSTO add-in via registry

Write-Host "=== PowerPoint AltTextManager Installation ===" -ForegroundColor Cyan
Write-Host ""

# Check for VSTO and DLL files
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$vstoPath = Join-Path $scriptDir "bin\Debug\AltTextManager.vsto"
$dllPath = Join-Path $scriptDir "bin\Debug\AltTextManager.dll"

if (-not (Test-Path $vstoPath)) {
    Write-Host "Error: VSTO file not found: $vstoPath" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please build the project first:" -ForegroundColor Yellow
    Write-Host "  .\build.ps1" -ForegroundColor Cyan
    exit 1
}

if (-not (Test-Path $dllPath)) {
    Write-Host "Error: DLL file not found: $dllPath" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please build the project first:" -ForegroundColor Yellow
    Write-Host "  .\build.ps1" -ForegroundColor Cyan
    exit 1
}

Write-Host "Found VSTO: $vstoPath" -ForegroundColor Green
Write-Host "Found DLL: $dllPath" -ForegroundColor Green
Write-Host ""

# Registry path for PowerPoint add-in
$registryPath = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\AltTextManager"

Write-Host "Registering add-in in registry..." -ForegroundColor Yellow

try {
    # Create registry key if not exists
    if (-not (Test-Path $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
        Write-Host "Created registry key: $registryPath" -ForegroundColor Green
    } else {
        Write-Host "Registry key exists: $registryPath" -ForegroundColor Green
    }

    # Set Manifest path
    Set-ItemProperty -Path $registryPath -Name "Manifest" -Value "$vstoPath|vstolocal" -Type String
    Write-Host "Set Manifest property" -ForegroundColor Green

    # Set FriendlyName
    Set-ItemProperty -Path $registryPath -Name "FriendlyName" -Value "PowerPoint AltTextManager" -Type String
    Write-Host "Set FriendlyName property" -ForegroundColor Green

    # Set Description
    Set-ItemProperty -Path $registryPath -Name "Description" -Value "Manage alternative text for PowerPoint shapes" -Type String
    Write-Host "Set Description property" -ForegroundColor Green

    # Set LoadBehavior (3 = load on startup)
    Set-ItemProperty -Path $registryPath -Name "LoadBehavior" -Value 3 -Type DWord
    Write-Host "Set LoadBehavior property (value: 3)" -ForegroundColor Green

    Write-Host ""
    Write-Host "Installation completed successfully!" -ForegroundColor Green
    Write-Host ""

    # Display installed values
    Write-Host "Installed registry values:" -ForegroundColor Cyan
    Get-ItemProperty -Path $registryPath | Format-List Manifest, FriendlyName, Description, LoadBehavior

    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Cyan
    Write-Host "1. Open PowerPoint" -ForegroundColor Gray
    Write-Host "2. Go to File > Options > Add-ins" -ForegroundColor Gray
    Write-Host "3. Select 'Manage: COM Add-ins' and click 'Go'" -ForegroundColor Gray
    Write-Host "4. Check if 'PowerPoint AltTextManager' appears in the list" -ForegroundColor Gray
    Write-Host ""

    Write-Host "Troubleshooting:" -ForegroundColor Yellow
    Write-Host "If add-in doesn't load, check PowerPoint Trust Center settings:" -ForegroundColor Gray
    Write-Host "  File > Options > Trust Center > Trust Center Settings" -ForegroundColor Gray
    Write-Host "  Add-ins > Uncheck 'Require Application Add-ins to be signed'" -ForegroundColor Gray

} catch {
    Write-Host ""
    Write-Host "Installation error!" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}
