# PowerShell script to uninstall Outlook Add-in from Registry (Classic Outlook)
# Run as Administrator

param(
    [string]$AddinName = "EmailIDViewer",
    [string]$OfficeVersion = "16.0"
)

Write-Host "=== Classic Outlook Add-in Registry Uninstaller ===" -ForegroundColor Red
Write-Host ""

Write-Host "Office version: $OfficeVersion" -ForegroundColor Cyan
Write-Host "Add-in name: $AddinName" -ForegroundColor Cyan
Write-Host ""

try {
    # Registry path for Office add-ins
    $registryPath = "HKCU:\SOFTWARE\Microsoft\Office\$OfficeVersion\WEF\Developer"
    
    # Check if the registry path exists
    if (!(Test-Path $registryPath)) {
        Write-Warning "Registry path not found: $registryPath"
        Write-Host "The add-in may not be installed or the Office version is incorrect."
        exit 1
    }
    
    # Check if the add-in entry exists
    $addinEntry = Get-ItemProperty -Path $registryPath -Name $AddinName -ErrorAction SilentlyContinue
    
    if (!$addinEntry) {
        Write-Warning "Add-in '$AddinName' not found in registry."
        Write-Host "The add-in may already be uninstalled."
        exit 1
    }
    
    Write-Host "Found add-in entry: $($addinEntry.$AddinName)" -ForegroundColor Yellow
    
    $confirm = Read-Host "Remove this add-in from registry? (y/n)"
    if ($confirm -eq "y") {
        # Remove the add-in entry
        Remove-ItemProperty -Path $registryPath -Name $AddinName
        
        Write-Host ""
        Write-Host "✅ Add-in successfully uninstalled!" -ForegroundColor Green
        Write-Host ""
        Write-Host "Next steps:" -ForegroundColor Cyan
        Write-Host "1. Restart Outlook completely (close all Outlook windows)" -ForegroundColor White
        Write-Host "2. The add-in should no longer appear in Outlook" -ForegroundColor White
    } else {
        Write-Host "Uninstall cancelled." -ForegroundColor Yellow
    }
    
} catch {
    Write-Error "Failed to uninstall add-in: $($_.Exception.Message)"
}

Write-Host ""
Read-Host "Press Enter to exit"
