#Requires -Version 5.1
<#
.SYNOPSIS
    Pre-flight backup for Outlook Favorites migration
.DESCRIPTION
    Creates a comprehensive backup of Outlook Favorites and related configuration
    before any migration operations. Run this on BOTH machines before importing.
.PARAMETER BackupFolder
    Destination folder for backups. Defaults to Desktop\OutlookBackup_[timestamp]
.EXAMPLE
    .\OutlookFavorites_Preflight.ps1
    .\OutlookFavorites_Preflight.ps1 -BackupFolder "D:\Backups\OutlookConfig"
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$BackupFolder
)

# Generate timestamped backup folder if not specified
if (-not $BackupFolder) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $BackupFolder = "$env:USERPROFILE\Desktop\OutlookBackup_$timestamp"
}

# Create backup folder
New-Item -ItemType Directory -Path $BackupFolder -Force | Out-Null

$script:LogPath = Join-Path $BackupFolder "backup_log.txt"

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        "INFO"    { Write-Host $logMessage -ForegroundColor Cyan }
        "WARN"    { Write-Host $logMessage -ForegroundColor Yellow }
        "ERROR"   { Write-Host $logMessage -ForegroundColor Red }
        "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
    }
    
    Add-Content -Path $script:LogPath -Value $logMessage
}

function Test-OutlookRunning {
    $outlook = Get-Process outlook -ErrorAction SilentlyContinue
    return $null -ne $outlook
}

# ============================================================
# START BACKUP
# ============================================================

Write-Host ""
Write-Host "=========================================" -ForegroundColor Magenta
Write-Host "  Outlook Favorites Pre-Flight Backup" -ForegroundColor Magenta
Write-Host "=========================================" -ForegroundColor Magenta
Write-Host ""

Write-Log "Backup destination: $BackupFolder"
Write-Log "Computer: $env:COMPUTERNAME"
Write-Log "User: $env:USERNAME"
Write-Log "Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

$backupResults = @{
    Favorites = $false
    Registry = $false
    NavPaneXML = $false
    ProfileInfo = $false
}

# ============================================================
# 1. EXPORT FAVORITES VIA COM (if Outlook is open)
# ============================================================

Write-Host ""
Write-Log "--- Step 1: Favorites Export (via Outlook COM) ---"

if (-not (Test-OutlookRunning)) {
    Write-Log "Outlook is not running - skipping COM-based favorites export" -Level WARN
    Write-Log "To include favorites export, open Outlook and run this script again" -Level WARN
}
else {
    try {
        $outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
        $namespace = $outlook.GetNamespace("MAPI")
        $explorer = $outlook.ActiveExplorer()
        
        if ($null -eq $explorer) {
            Write-Log "No active Outlook window found - skipping COM export" -Level WARN
        }
        else {
            $navPane = $explorer.NavigationPane
            $mailModule = $navPane.Modules.Item("Mail")
            $favoritesGroup = $mailModule.NavigationGroups.Item(1)
            
            $favorites = @()
            for ($i = 1; $i -le $favoritesGroup.NavigationFolders.Count; $i++) {
                try {
                    $navFolder = $favoritesGroup.NavigationFolders.Item($i)
                    $favorites += @{
                        Name = $navFolder.Folder.Name
                        Path = $navFolder.Folder.FolderPath
                        Position = $i
                    }
                }
                catch {
                    Write-Log "Could not read favorite at position $i" -Level WARN
                }
            }
            
            $exportData = @{
                ExportDate = (Get-Date).ToString("o")
                ComputerName = $env:COMPUTERNAME
                UserName = $env:USERNAME
                OutlookVersion = $outlook.Version
                FavoritesCount = $favorites.Count
                Favorites = $favorites
            }
            
            $favoritesFile = Join-Path $BackupFolder "Favorites.json"
            $exportData | ConvertTo-Json -Depth 10 | Out-File -FilePath $favoritesFile -Encoding UTF8
            
            Write-Log "Exported $($favorites.Count) favorites to Favorites.json" -Level SUCCESS
            $backupResults.Favorites = $true
            
            # Also save a human-readable list
            $readableFile = Join-Path $BackupFolder "Favorites_ReadableList.txt"
            $readableContent = @"
Outlook Favorites Backup
========================
Exported: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Computer: $env:COMPUTERNAME
Total Favorites: $($favorites.Count)

Favorites List:
---------------
"@
            foreach ($fav in $favorites) {
                $readableContent += "`n$($fav.Position). $($fav.Name)`n   Path: $($fav.Path)`n"
            }
            $readableContent | Out-File -FilePath $readableFile -Encoding UTF8
            Write-Log "Created human-readable list: Favorites_ReadableList.txt" -Level SUCCESS
        }
    }
    catch {
        Write-Log "Failed to export favorites via COM: $_" -Level ERROR
    }
}

# ============================================================
# 2. BACKUP REGISTRY KEYS
# ============================================================

Write-Host ""
Write-Log "--- Step 2: Registry Backup ---"

# Outlook profiles
$regProfilesPath = "HKCU\Software\Microsoft\Office\16.0\Outlook\Profiles"
$regProfilesFile = Join-Path $BackupFolder "Registry_OutlookProfiles.reg"

try {
    $regCheck = reg query $regProfilesPath 2>&1
    if ($LASTEXITCODE -eq 0) {
        reg export $regProfilesPath $regProfilesFile /y 2>&1 | Out-Null
        Write-Log "Exported Outlook Profiles registry: Registry_OutlookProfiles.reg" -Level SUCCESS
        $backupResults.Registry = $true
    }
    else {
        # Try Office 15.0 (Outlook 2013)
        $regProfilesPath = "HKCU\Software\Microsoft\Office\15.0\Outlook\Profiles"
        $regCheck = reg query $regProfilesPath 2>&1
        if ($LASTEXITCODE -eq 0) {
            reg export $regProfilesPath $regProfilesFile /y 2>&1 | Out-Null
            Write-Log "Exported Outlook 2013 Profiles registry: Registry_OutlookProfiles.reg" -Level SUCCESS
            $backupResults.Registry = $true
        }
        else {
            Write-Log "Could not find Outlook profiles registry key" -Level WARN
        }
    }
}
catch {
    Write-Log "Failed to export registry: $_" -Level ERROR
}

# Also backup general Outlook settings
$regOutlookPath = "HKCU\Software\Microsoft\Office\16.0\Outlook"
$regOutlookFile = Join-Path $BackupFolder "Registry_OutlookSettings.reg"

try {
    reg export $regOutlookPath $regOutlookFile /y 2>&1 | Out-Null
    Write-Log "Exported Outlook settings registry: Registry_OutlookSettings.reg" -Level SUCCESS
}
catch {
    Write-Log "Failed to export Outlook settings registry" -Level WARN
}

# ============================================================
# 3. BACKUP NAVIGATION PANE XML FILES
# ============================================================

Write-Host ""
Write-Log "--- Step 3: Navigation Pane XML Backup ---"

$outlookLocalAppData = "$env:LOCALAPPDATA\Microsoft\Outlook"
$xmlFiles = Get-ChildItem -Path $outlookLocalAppData -Filter "*.xml" -ErrorAction SilentlyContinue

if ($xmlFiles.Count -gt 0) {
    $xmlBackupFolder = Join-Path $BackupFolder "NavPaneXML"
    New-Item -ItemType Directory -Path $xmlBackupFolder -Force | Out-Null
    
    foreach ($xml in $xmlFiles) {
        try {
            Copy-Item $xml.FullName -Destination $xmlBackupFolder -Force
            Write-Log "Copied: $($xml.Name)" -Level SUCCESS
        }
        catch {
            Write-Log "Failed to copy $($xml.Name): $_" -Level WARN
        }
    }
    $backupResults.NavPaneXML = $true
}
else {
    Write-Log "No Navigation Pane XML files found in $outlookLocalAppData" -Level WARN
}

# ============================================================
# 4. CAPTURE PROFILE & ACCOUNT INFO
# ============================================================

Write-Host ""
Write-Log "--- Step 4: Profile Information ---"

$profileInfo = @{
    CaptureDate = (Get-Date).ToString("o")
    ComputerName = $env:COMPUTERNAME
    UserName = $env:USERNAME
    WindowsVersion = (Get-CimInstance Win32_OperatingSystem).Caption
    OutlookDataFiles = @()
    OutlookProcessInfo = $null
}

# Get Outlook data files (PST/OST locations)
$dataFilePaths = @(
    "$env:LOCALAPPDATA\Microsoft\Outlook",
    "$env:USERPROFILE\Documents\Outlook Files"
)

foreach ($path in $dataFilePaths) {
    if (Test-Path $path) {
        $files = Get-ChildItem -Path $path -Include "*.pst","*.ost" -Recurse -ErrorAction SilentlyContinue
        foreach ($file in $files) {
            $profileInfo.OutlookDataFiles += @{
                Name = $file.Name
                Path = $file.FullName
                SizeGB = [math]::Round($file.Length / 1GB, 2)
                LastModified = $file.LastWriteTime.ToString("o")
            }
        }
    }
}

# Get Outlook process info if running
$outlookProcess = Get-Process outlook -ErrorAction SilentlyContinue
if ($outlookProcess) {
    $profileInfo.OutlookProcessInfo = @{
        Path = $outlookProcess.Path
        Version = $outlookProcess.FileVersion
        StartTime = $outlookProcess.StartTime.ToString("o")
    }
}

$profileInfoFile = Join-Path $BackupFolder "ProfileInfo.json"
$profileInfo | ConvertTo-Json -Depth 10 | Out-File -FilePath $profileInfoFile -Encoding UTF8
Write-Log "Saved profile information: ProfileInfo.json" -Level SUCCESS
$backupResults.ProfileInfo = $true

# ============================================================
# 5. CREATE RESTORE INSTRUCTIONS
# ============================================================

Write-Host ""
Write-Log "--- Creating Restore Instructions ---"

$restoreInstructions = @"
OUTLOOK FAVORITES BACKUP - RESTORE INSTRUCTIONS
================================================
Backup Created: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Computer: $env:COMPUTERNAME

CONTENTS OF THIS BACKUP:
------------------------
$(if ($backupResults.Favorites) { "[X]" } else { "[ ]" }) Favorites.json - Favorites data for import script
$(if ($backupResults.Favorites) { "[X]" } else { "[ ]" }) Favorites_ReadableList.txt - Human-readable list (for manual rebuild)
$(if ($backupResults.Registry) { "[X]" } else { "[ ]" }) Registry_OutlookProfiles.reg - Outlook profile registry keys
$(if ($backupResults.Registry) { "[X]" } else { "[ ]" }) Registry_OutlookSettings.reg - Outlook settings registry keys
$(if ($backupResults.NavPaneXML) { "[X]" } else { "[ ]" }) NavPaneXML\ - Navigation Pane XML configuration files
$(if ($backupResults.ProfileInfo) { "[X]" } else { "[ ]" }) ProfileInfo.json - System and profile information

HOW TO RESTORE:
---------------

Option 1: Use the Import Script (Recommended)
   Copy Favorites.json to the new PC and use OutlookFavoritesManager.ps1:
   
   .\OutlookFavoritesManager.ps1 -Action Import -FilePath "path\to\Favorites.json"

Option 2: Manual Registry Restore (if things go wrong)
   1. Close Outlook completely
   2. Double-click Registry_OutlookProfiles.reg and confirm the import
   3. Copy files from NavPaneXML\ to %LOCALAPPDATA%\Microsoft\Outlook\
   4. Reopen Outlook

Option 3: Manual Rebuild (last resort)
   Open Favorites_ReadableList.txt and manually re-add each folder to Favorites

NOTES:
------
- The Favorites.json file is what you need for the import script
- Keep this entire backup folder until migration is confirmed successful
- If Outlook wasn't running during backup, re-run with Outlook open to capture Favorites.json

"@

$instructionsFile = Join-Path $BackupFolder "RESTORE_INSTRUCTIONS.txt"
$restoreInstructions | Out-File -FilePath $instructionsFile -Encoding UTF8
Write-Log "Created: RESTORE_INSTRUCTIONS.txt" -Level SUCCESS

# ============================================================
# SUMMARY
# ============================================================

Write-Host ""
Write-Host "=========================================" -ForegroundColor Magenta
Write-Host "  BACKUP COMPLETE" -ForegroundColor Magenta
Write-Host "=========================================" -ForegroundColor Magenta
Write-Host ""

Write-Log "Backup location: $BackupFolder" -Level SUCCESS

Write-Host ""
Write-Host "Backup Contents:" -ForegroundColor White
Write-Host "  Favorites export:    $(if ($backupResults.Favorites) { 'YES' } else { 'NO (Outlook not running)' })" -ForegroundColor $(if ($backupResults.Favorites) { 'Green' } else { 'Yellow' })
Write-Host "  Registry backup:     $(if ($backupResults.Registry) { 'YES' } else { 'NO' })" -ForegroundColor $(if ($backupResults.Registry) { 'Green' } else { 'Yellow' })
Write-Host "  NavPane XML backup:  $(if ($backupResults.NavPaneXML) { 'YES' } else { 'NO (none found)' })" -ForegroundColor $(if ($backupResults.NavPaneXML) { 'Green' } else { 'Yellow' })
Write-Host "  Profile info:        $(if ($backupResults.ProfileInfo) { 'YES' } else { 'NO' })" -ForegroundColor $(if ($backupResults.ProfileInfo) { 'Green' } else { 'Yellow' })

Write-Host ""
if (-not $backupResults.Favorites) {
    Write-Host "NOTE: Open Outlook and run this script again to capture Favorites.json" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Next steps:" -ForegroundColor White
Write-Host "  1. Copy the entire '$BackupFolder' folder to a USB drive or cloud storage" -ForegroundColor Gray
Write-Host "  2. On the new PC, run this same script BEFORE importing (for a baseline)" -ForegroundColor Gray
Write-Host "  3. Then run: .\OutlookFavoritesManager.ps1 -Action Import -FilePath 'path\to\Favorites.json'" -ForegroundColor Gray
Write-Host ""

# Open backup folder in Explorer
Write-Host "Opening backup folder..." -ForegroundColor Cyan
Start-Process explorer.exe -ArgumentList $BackupFolder
