# List of directories to search (no admin rights required)
$searchPaths = @(
    "$env:ProgramFiles",
    "$env:ProgramFiles(x86)",
    "$env:LOCALAPPDATA",
    "$env:APPDATA"
)

# Optional: add any known Java-heavy app paths
$searchPaths += "$env:USERPROFILE\Documents", "$env:USERPROFILE\Downloads"

$results = @()

foreach ($basePath in $searchPaths) {
    if (Test-Path $basePath) {
        # Recursively find all java.exe files, skipping errors
        Get-ChildItem -Path $basePath -Recurse -Filter "java.exe" -ErrorAction SilentlyContinue -Force |
        ForEach-Object {
            $javaPath = $_.FullName
            try {
                $versionOutput = & "$javaPath" -version 2>&1 | Select-String 'version' | Select-Object -First 1
                if ($versionOutput) {
                    $results += [PSCustomObject]@{
                        JavaPath      = $javaPath
                        JavaVersion   = $versionOutput -replace 'java version|openjdk version', '' -replace '"', '' -replace '^\s+', ''
                        ParentAppHint = ($javaPath -replace '\\bin\\java.exe$', '') -replace '^.*\\', ''
                    }
                }
            } catch {
                # Skip errors (inaccessible or broken binaries)
            }
        }
    }
}

# Output results as a table
if ($results.Count -eq 0) {
    Write-Output "No java.exe binaries found in user-accessible folders."
} else {
    $results | Sort-Object JavaPath | Format-Table -AutoSize
}