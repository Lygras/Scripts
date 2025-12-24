#Requires -Version 5.1
<#
.SYNOPSIS
    Outlook Favorites Manager - Export, Import, and Rollback favorites
.DESCRIPTION
    Manages Outlook Navigation Pane Favorites across machines.
    - Export: Saves current favorites to a JSON file with metadata
    - Import: Restores favorites from export file, logs what was added
    - Rollback: Removes favorites that were added during the last import
.PARAMETER Action
    The action to perform: Export, Import, or Rollback
.PARAMETER FilePath
    Path to the favorites file. Defaults to Desktop\OutlookFavorites.json
.EXAMPLE
    .\OutlookFavoritesManager.ps1 -Action Export
    .\OutlookFavoritesManager.ps1 -Action Import
    .\OutlookFavoritesManager.ps1 -Action Rollback
#>

param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("Export", "Import", "Rollback")]
    [string]$Action,
    
    [Parameter(Mandatory=$false)]
    [string]$FilePath = "$env:USERPROFILE\Desktop\OutlookFavorites.json"
)

# Configuration
$script:LogPath = "$env:USERPROFILE\Desktop\OutlookFavorites_Log.txt"
$script:RollbackPath = "$env:USERPROFILE\Desktop\OutlookFavorites_Rollback.json"

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Console output with colors
    switch ($Level) {
        "INFO"    { Write-Host $logMessage -ForegroundColor Cyan }
        "WARN"    { Write-Host $logMessage -ForegroundColor Yellow }
        "ERROR"   { Write-Host $logMessage -ForegroundColor Red }
        "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
    }
    
    # Append to log file
    Add-Content -Path $script:LogPath -Value $logMessage
}

function Get-OutlookObjects {
    <#
    .DESCRIPTION
        Initializes Outlook COM objects and returns them.
        Validates that Outlook is running and accessible.
    #>
    
    Write-Log "Connecting to Outlook..."
    
    try {
        $outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
    }
    catch {
        Write-Log "Failed to connect to Outlook. Make sure Outlook is running." -Level ERROR
        Write-Log "Error: $_" -Level ERROR
        return $null
    }
    
    try {
        $namespace = $outlook.GetNamespace("MAPI")
        $explorer = $outlook.ActiveExplorer()
        
        if ($null -eq $explorer) {
            Write-Log "No active Outlook Explorer window. Please open Outlook and try again." -Level ERROR
            return $null
        }
        
        $navPane = $explorer.NavigationPane
        $mailModule = $navPane.Modules.Item("Mail")
        $favoritesGroup = $mailModule.NavigationGroups.Item(1)
        
        Write-Log "Successfully connected to Outlook" -Level SUCCESS
        
        return @{
            Outlook = $outlook
            Namespace = $namespace
            FavoritesGroup = $favoritesGroup
        }
    }
    catch {
        Write-Log "Failed to access Outlook navigation pane: $_" -Level ERROR
        return $null
    }
}

function Get-CurrentFavorites {
    param($FavoritesGroup)
    
    $favorites = @()
    
    for ($i = 1; $i -le $FavoritesGroup.NavigationFolders.Count; $i++) {
        try {
            $navFolder = $FavoritesGroup.NavigationFolders.Item($i)
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
    
    return $favorites
}

function Export-Favorites {
    Write-Log "========== EXPORT STARTED ==========" -Level INFO
    
    $outlookObjects = Get-OutlookObjects
    if ($null -eq $outlookObjects) {
        return $false
    }
    
    $favoritesGroup = $outlookObjects.FavoritesGroup
    
    Write-Log "Reading current favorites..."
    $favorites = Get-CurrentFavorites -FavoritesGroup $favoritesGroup
    
    if ($favorites.Count -eq 0) {
        Write-Log "No favorites found to export" -Level WARN
        return $false
    }
    
    # Build export object with metadata
    $exportData = @{
        ExportDate = (Get-Date).ToString("o")
        ComputerName = $env:COMPUTERNAME
        UserName = $env:USERNAME
        OutlookVersion = $outlookObjects.Outlook.Version
        FavoritesCount = $favorites.Count
        Favorites = $favorites
    }
    
    try {
        $exportData | ConvertTo-Json -Depth 10 | Out-File -FilePath $FilePath -Encoding UTF8
        Write-Log "Exported $($favorites.Count) favorites to: $FilePath" -Level SUCCESS
        
        Write-Log "--- Exported Favorites ---" -Level INFO
        foreach ($fav in $favorites) {
            Write-Log "  [$($fav.Position)] $($fav.Name): $($fav.Path)" -Level INFO
        }
        
        return $true
    }
    catch {
        Write-Log "Failed to write export file: $_" -Level ERROR
        return $false
    }
}

function Import-Favorites {
    Write-Log "========== IMPORT STARTED ==========" -Level INFO
    
    # Validate import file exists
    if (-not (Test-Path $FilePath)) {
        Write-Log "Import file not found: $FilePath" -Level ERROR
        return $false
    }
    
    # Read import file
    try {
        $importData = Get-Content $FilePath -Raw | ConvertFrom-Json
        Write-Log "Loaded import file from: $FilePath" -Level INFO
        Write-Log "Export was from $($importData.ComputerName) on $($importData.ExportDate)" -Level INFO
        Write-Log "Contains $($importData.FavoritesCount) favorites" -Level INFO
    }
    catch {
        Write-Log "Failed to read import file: $_" -Level ERROR
        return $false
    }
    
    $outlookObjects = Get-OutlookObjects
    if ($null -eq $outlookObjects) {
        return $false
    }
    
    $favoritesGroup = $outlookObjects.FavoritesGroup
    $namespace = $outlookObjects.Namespace
    
    # Get current favorites to avoid duplicates
    $currentFavorites = Get-CurrentFavorites -FavoritesGroup $favoritesGroup
    $currentPaths = $currentFavorites | ForEach-Object { $_.Path }
    
    # Track what we add for rollback
    $addedFavorites = @()
    $skippedFavorites = @()
    $failedFavorites = @()
    
    Write-Log "--- Processing Import ---" -Level INFO
    
    foreach ($fav in $importData.Favorites) {
        $path = $fav.Path
        $name = $fav.Name
        
        # Check if already in favorites
        if ($currentPaths -contains $path) {
            Write-Log "SKIP (already exists): $name" -Level WARN
            $skippedFavorites += $fav
            continue
        }
        
        # Try to find and add the folder
        try {
            $folder = $namespace.GetFolderFromPath($path)
            $favoritesGroup.NavigationFolders.Add($folder) | Out-Null
            Write-Log "ADDED: $name ($path)" -Level SUCCESS
            $addedFavorites += @{
                Name = $name
                Path = $path
                ImportedAt = (Get-Date).ToString("o")
            }
        }
        catch {
            Write-Log "FAILED: $name ($path) - $_" -Level ERROR
            $failedFavorites += @{
                Name = $name
                Path = $path
                Error = $_.ToString()
            }
        }
    }
    
    # Save rollback data
    $rollbackData = @{
        ImportDate = (Get-Date).ToString("o")
        SourceFile = $FilePath
        ComputerName = $env:COMPUTERNAME
        AddedFavorites = $addedFavorites
    }
    
    try {
        $rollbackData | ConvertTo-Json -Depth 10 | Out-File -FilePath $script:RollbackPath -Encoding UTF8
        Write-Log "Rollback data saved to: $($script:RollbackPath)" -Level INFO
    }
    catch {
        Write-Log "Warning: Could not save rollback data: $_" -Level WARN
    }
    
    # Summary
    Write-Log "========== IMPORT SUMMARY ==========" -Level INFO
    Write-Log "Successfully added: $($addedFavorites.Count)" -Level SUCCESS
    Write-Log "Skipped (duplicates): $($skippedFavorites.Count)" -Level WARN
    Write-Log "Failed: $($failedFavorites.Count)" -Level $(if ($failedFavorites.Count -gt 0) { "ERROR" } else { "INFO" })
    
    if ($failedFavorites.Count -gt 0) {
        Write-Log "--- Failed Items (folder may not exist on this machine) ---" -Level ERROR
        foreach ($fail in $failedFavorites) {
            Write-Log "  $($fail.Name): $($fail.Path)" -Level ERROR
        }
    }
    
    return $true
}

function Invoke-Rollback {
    Write-Log "========== ROLLBACK STARTED ==========" -Level INFO
    
    # Check for rollback file
    if (-not (Test-Path $script:RollbackPath)) {
        Write-Log "No rollback data found at: $($script:RollbackPath)" -Level ERROR
        Write-Log "Rollback is only available after an import operation." -Level INFO
        return $false
    }
    
    # Read rollback data
    try {
        $rollbackData = Get-Content $script:RollbackPath -Raw | ConvertFrom-Json
        Write-Log "Found rollback data from import on $($rollbackData.ImportDate)" -Level INFO
    }
    catch {
        Write-Log "Failed to read rollback file: $_" -Level ERROR
        return $false
    }
    
    if ($rollbackData.AddedFavorites.Count -eq 0) {
        Write-Log "No favorites were added in the last import - nothing to roll back" -Level WARN
        return $true
    }
    
    $outlookObjects = Get-OutlookObjects
    if ($null -eq $outlookObjects) {
        return $false
    }
    
    $favoritesGroup = $outlookObjects.FavoritesGroup
    
    $removedCount = 0
    $notFoundCount = 0
    
    Write-Log "--- Processing Rollback ---" -Level INFO
    
    foreach ($fav in $rollbackData.AddedFavorites) {
        $path = $fav.Path
        $name = $fav.Name
        $found = $false
        
        # Find and remove the navigation folder
        # Note: We iterate backwards to avoid index shifting issues
        for ($i = $favoritesGroup.NavigationFolders.Count; $i -ge 1; $i--) {
            try {
                $navFolder = $favoritesGroup.NavigationFolders.Item($i)
                if ($navFolder.Folder.FolderPath -eq $path) {
                    $favoritesGroup.NavigationFolders.Remove($navFolder)
                    Write-Log "REMOVED: $name" -Level SUCCESS
                    $removedCount++
                    $found = $true
                    break
                }
            }
            catch {
                # Folder might have been manually removed
                continue
            }
        }
        
        if (-not $found) {
            Write-Log "NOT FOUND (may have been manually removed): $name" -Level WARN
            $notFoundCount++
        }
    }
    
    # Remove rollback file after successful rollback
    try {
        Remove-Item $script:RollbackPath -Force
        Write-Log "Rollback data file removed" -Level INFO
    }
    catch {
        Write-Log "Could not remove rollback file: $_" -Level WARN
    }
    
    # Summary
    Write-Log "========== ROLLBACK SUMMARY ==========" -Level INFO
    Write-Log "Removed: $removedCount" -Level SUCCESS
    Write-Log "Not found: $notFoundCount" -Level $(if ($notFoundCount -gt 0) { "WARN" } else { "INFO" })
    
    return $true
}

# Main execution
Write-Log "Outlook Favorites Manager v1.0" -Level INFO
Write-Log "Action: $Action" -Level INFO

switch ($Action) {
    "Export"   { $result = Export-Favorites }
    "Import"   { $result = Import-Favorites }
    "Rollback" { $result = Invoke-Rollback }
}

if ($result) {
    Write-Log "Operation completed successfully" -Level SUCCESS
}
else {
    Write-Log "Operation failed - check log for details" -Level ERROR
}

Write-Log "Log file: $($script:LogPath)" -Level INFO
