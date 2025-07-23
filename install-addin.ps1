# PowerShell script to install Outlook Add-in via Registry (Classic Outlook)
# Run as Administrator

param(
    [string]$ManifestPath = "",
    [string]$AddinName = "EmailIDViewer",
    [string]$OfficeVersion = "16.0"
)

Write-Host "=== Classic Outlook Add-in Registry Installer ===" -ForegroundColor Green
Write-Host ""

# Get the manifest path if not provided
if ([string]::IsNullOrEmpty($ManifestPath)) {
    $ManifestPath = Join-Path $PSScriptRoot "manifest.xml"
}

# Verify manifest file exists
if (!(Test-Path $ManifestPath)) {
    Write-Error "Manifest file not found at: $ManifestPath"
    Write-Host "Please ensure the manifest.xml file exists in the project folder."
    exit 1
}

$ManifestPath = Resolve-Path $ManifestPath

Write-Host "Manifest file: $ManifestPath" -ForegroundColor Cyan
Write-Host "Office version: $OfficeVersion" -ForegroundColor Cyan
Write-Host "Add-in name: $AddinName" -ForegroundColor Cyan
Write-Host ""

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

if (!$isAdmin) {
    Write-Warning "This script should be run as Administrator for registry modifications."
    $continue = Read-Host "Continue anyway? (y/n)"
    if ($continue -ne "y") {
        exit 1
    }
}

try {
    # Registry path for Office add-ins
    $registryPath = "HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\WEF\Developer"
    
    Write-Host "Creating registry path if it doesn't exist..." -ForegroundColor Yellow
    
    # Ensure the registry path exists
    if (!(Test-Path $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
        Write-Host "Created registry path: $registryPath" -ForegroundColor Green
    }
    
    # Add the manifest path to registry
    Write-Host "Adding add-in to registry..." -ForegroundColor Yellow
    Set-ItemProperty -Path $registryPath -Name $AddinName -Value $ManifestPath
    
    Write-Host ""
    Write-Host "✅ Add-in successfully registered!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Cyan
    Write-Host "1. Start your development server: npm start" -ForegroundColor White
    Write-Host "2. Restart Outlook completely (close all Outlook windows)" -ForegroundColor White
    Write-Host "3. Open Outlook and open an email to see the add-in" -ForegroundColor White
    Write-Host ""
    Write-Host "To uninstall later, run:" -ForegroundColor Yellow
    Write-Host "Remove-ItemProperty -Path '$registryPath' -Name '$AddinName'" -ForegroundColor Gray
    
} catch {
    Write-Error "Failed to register add-in: $($_.Exception.Message)"
    Write-Host ""
    Write-Host "Manual registration steps:" -ForegroundColor Yellow
    Write-Host "1. Open Registry Editor (regedit) as Administrator"
    Write-Host "2. Navigate to: HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\$OfficeVersion\WEF\Developer"
    Write-Host "3. Create a new String Value named '$AddinName'"
    Write-Host "4. Set its value to: $ManifestPath"
    Write-Host "5. Restart Outlook"
}

Write-Host ""
Read-Host "Press Enter to exit"
