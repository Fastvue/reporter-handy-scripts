  <#
.SYNOPSIS
    Deletes Fastvue Syslog archive files older than N days.

.DESCRIPTION

    Change the $ArchiveRootPath and the $DaysToKeep variables in the User-configurable settings section as desired.

    Supports filenames inside each device/host folder with the following formats:
      {hostname}-YYYY-MM-dd.log.zip
      {hostname}-YYYY-MM-dd-HH.log.zip

    For each archive "pair" (basename):
      - Parses date from the basename
      - If older than $DaysToKeep:
          Deletes any of:
            - .log.zip (if present)
            - .log.sha256 (if present)

    Logs all actions to ArchiveCleanupLogs\ArchiveCleanup_*.log
    in the same folder as this script.
#> 


########################
# User-configurable settings
########################

# Root archive folder (contains the IP/hostname subfolders)
$ArchiveRootPath = 'C:\ProgramData\Fastvue\Syslog Server\Archive'

# Age threshold in days (e.g. 365 = delete archives older than 1 year)
$DaysToKeep = 365

########################
# Advanced (usually no need to change)
########################

# Matches:
#   prefix-YYYY-MM-dd.log
#   prefix-YYYY-MM-dd-HH.log
# (prefix can be any hostname or IP)
$FileNameRegex = '^(?<prefix>.+)-(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})(-(?<hour>\d{2}))?\.log$'
$FileNameRegexObj = [regex]::new($FileNameRegex)

########################
# Script setup
########################

if ($PSScriptRoot -and ($PSScriptRoot -ne '')) {
    $ScriptDirectory = $PSScriptRoot
} elseif ($MyInvocation.MyCommand.Path) {
    $ScriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    $ScriptDirectory = (Get-Location).Path
}

$LogDirectory = Join-Path $ScriptDirectory 'ArchiveCleanupLogs'
if (-not (Test-Path $LogDirectory)) {
    New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
}

$LogFile = Join-Path $LogDirectory ("ArchiveCleanup_{0:yyyy-MM-dd_HH-mm-ss}.log" -f (Get-Date))

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[{0}] [{1}] {2}" -f $timestamp, $Level, $Message

    Write-Host $line
    Add-Content -Path $LogFile -Value $line
}

########################
# Main logic
########################

Write-Log "Starting Fastvue Syslog archive cleanup (single-pass)."
Write-Log "Archive root: $ArchiveRootPath"
Write-Log "Days to keep: $DaysToKeep"

if (-not (Test-Path $ArchiveRootPath)) {
    Write-Log "Archive root path does not exist: $ArchiveRootPath" 'ERROR'
    throw "Archive root path does not exist."
}

$CutoffDate = (Get-Date).AddDays(-1 * $DaysToKeep)
Write-Log "Cutoff date/time (local): $($CutoffDate.ToString('yyyy-MM-dd HH:mm:ss'))"

$deletedZipCount = 0
$deletedShaCount = 0
$skippedCount    = 0
$errorCount      = 0

# One pass: get both .log.zip and .log.sha256
try {
    $allFiles = Get-ChildItem -Path $ArchiveRootPath -Recurse -File -ErrorAction Stop |
        Where-Object {
            $_.Name -like '*.log.zip' -or $_.Name -like '*.log.sha256'
        }
} catch {
    Write-Log "Error enumerating archive files: $($_.Exception.Message)" 'ERROR'
    $allFiles = @()
}

if (-not $allFiles) {
    Write-Log "No archive files (.log.zip / .log.sha256) found under: $ArchiveRootPath"
    Write-Log "Cleanup complete (nothing to do)."
    return
}

Write-Log "Found $($allFiles.Count) archive-related files to evaluate."

# Group by basename, e.g. "192.168.3.1-2019-08-16-13.log"
$groups = $allFiles | Group-Object -Property BaseName

foreach ($group in $groups) {
    $baseName = $group.Name        # e.g. 192.168.3.1-2019-08-16-13.log
    $files    = $group.Group

    # Expect files with this basename, possibly:
    #   *.log.zip
    #   *.log.sha256
    $zipFile = $files | Where-Object { $_.Name -like '*.log.zip' } | Select-Object -First 1
    $shaFile = $files | Where-Object { $_.Name -like '*.log.sha256' } | Select-Object -First 1

    # Parse date from basename
    $match = $FileNameRegexObj.Match($baseName)
    if (-not $match.Success) {
        Write-Log "Skipping files with unexpected base name format: $baseName" 'WARN'
        $skippedCount += $files.Count
        continue
    }

    try {
        $year  = [int]$match.Groups['year'].Value
        $month = [int]$match.Groups['month'].Value
        $day   = [int]$match.Groups['day'].Value

        $hour  = 0
        if ($match.Groups['hour'].Success -and $match.Groups['hour'].Value) {
            $hour = [int]$match.Groups['hour'].Value
        }

        $archiveDate = Get-Date -Year $year -Month $month -Day $day -Hour $hour -Minute 0 -Second 0
    }
    catch {
        Write-Log "Failed to parse date from base name: $baseName. Error: $($_.Exception.Message)" 'ERROR'
        $errorCount += $files.Count
        continue
    }

    # Only delete if archive date is strictly older than the cutoff
    if ($archiveDate -gt $CutoffDate) {
        $skippedCount += $files.Count
        continue
    }

    # Delete .log.zip if present
    if ($zipFile) {
        try {
            Remove-Item -Path $zipFile.FullName -Force -ErrorAction Stop
            Write-Log "Deleted archive file: $($zipFile.FullName)"
            $deletedZipCount++
        }
        catch {
            Write-Log "FAILED to delete archive file: $($zipFile.FullName). Error: $($_.Exception.Message)" 'ERROR'
            $errorCount++
        }
    }

    # Delete .log.sha256 if present (even if .zip was missing)
    if ($shaFile) {
        try {
            Remove-Item -Path $shaFile.FullName -Force -ErrorAction Stop
            Write-Log "Deleted checksum file: $($shaFile.FullName)"
            $deletedShaCount++
        }
        catch {
            Write-Log "FAILED to delete checksum file: $($shaFile.FullName). Error: $($_.Exception.Message)" 'ERROR'
            $errorCount++
        }
    }
}

########################
# Summary
########################

Write-Log "Cleanup complete."
Write-Log "Deleted .log.zip files   : $deletedZipCount"
Write-Log "Deleted .log.sha256 files: $deletedShaCount"
Write-Log "Skipped files            : $skippedCount"
Write-Log "Errors                   : $errorCount"
Write-Log "Log file written to: $LogFile"
