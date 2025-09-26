# This script will scan the registry for all keys associated with install,
# create a log file, and optionally delete them (not recommended - uninstall via Inno Setup unins000.exe)

$ProgID = "SolverWrapper"
$GuidPrefix = "71A00FE4-A1D4-47C3-BC7A-"
$LogPath = ".\COM_Registry_Log.txt"
$Delete = $false

$ProgIDPattern = "$ProgID.*"

$log = New-Object System.Collections.Generic.List[string]
function Log($msg) {
    $timestamped = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $msg"
    $log.Add($timestamped)
    Add-Content -Path $LogPath -Value $timestamped
    Write-Host $timestamped
}

function MatchKeys($basePath, $pattern) {
    Log "Scanning: $basePath with pattern: $pattern"
    try {
        $items = Get-ChildItem -Path "Registry::$basePath" -ErrorAction SilentlyContinue
        $matches = $items | Where-Object {
            $_.PSChildName -match "^$pattern" -or $_.PSChildName -match "^\{$pattern"
            } | ForEach-Object { $_.Name }
        Log "Matched $($matches.Count) keys"
        return $matches
    } catch {
        Log "Failed to scan: $basePath - $_"
        return @()
    }
}

function DeleteKey($keyPath) {
try {
   Remove-Item -Path "Registry::$keyPath" -Recurse -Force -ErrorAction Stop
   Log "Deleted: $keyPath"
} catch {
   Log "Failed to delete: $keyPath - $_"
}
}

# Registry hives to scan
$hives = @(
"HKCR",
"HKCR\WOW6432Node",
"HKCU\Software\Classes",
"HKCU\Software\Classes\Wow6432Node",
"HKCU\Software\Wow6432Node",
"HKLM\Software\Classes",
"HKLM\Software\Classes\WOW6432Node",
"HKLM\Software\WOW6432Node\Classes"
)

#"Software\Classes\TypeLib"

# Targets to match
$targets = @(
@{ Path = ""; Pattern = $ProgIDPattern },
@{ Path = "CLSID"; Pattern = "$GuidPrefix*" },
@{ Path = "Interface"; Pattern = "$GuidPrefix*" },
@{ Path = "TypeLib"; Pattern = "$GuidPrefix*" }
)

foreach ($hive in $hives) {
    foreach ($target in $targets) {
        $base = if ($target.Path -eq "") { $hive } else { "$hive\$($target.Path)" }
        $matches = MatchKeys $base $target.Pattern
        foreach ($match in $matches) {
            Log "Found: $match"
            if ($Delete) { DeleteKey $match }
        }
    }
}

# Detect Office version from Excel and Access
function Get-OfficeVersion($app) {
    $key = "HKCU:\Software\Microsoft\Office"
    $versions = Get-ChildItem -Path $key -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^\d+\.\d+$' } |
        Sort-Object -Property PSChildName -Descending

    foreach ($version in $versions) {
        $testPath = "$key\$($version.PSChildName)\$app\Security\Trusted Locations\$ProgID"
        if (Test-Path -Path $testPath) {
            return $version.PSChildName
        }
    }
    return $null
}

# Delete Trusted Location key
function FindTrustedLocation($app) {
    $version = Get-OfficeVersion $app
    if ($version) {
        $keyPath = "HKCU:\Software\Microsoft\Office\$version\$app\Security\Trusted Locations\$ProgID"
        if (Test-Path -Path $keyPath) {
            Log "Found Trusted Location: $keyPath"
            if ($Delete) {
                try {
                    Remove-Item -Path $keyPath -Recurse -Force -ErrorAction Stop
                    Log "Deleted Trusted Location: $keyPath"
                } catch {
                    Log "Failed to delete $keyPath - $_"
                }
            }
        } else {
            Log "No Trusted Location found for $app {$version}: $keyPath"
        }

    } else {
        Log "No Trusted Location found for $app"
    }
}

# Run cleanup
FindTrustedLocation "Excel"
FindTrustedLocation "Access"


# Save log to file
$log | Out-File -FilePath $LogPath -Encoding UTF8
Log "Log saved to: $LogPath"