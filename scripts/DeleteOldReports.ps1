# This script deletes Fastvue Reporter reports older than a configured age.
# It supports age thresholds in Days, Months, or Years.
#
# How to use:
# 1. Set $apiBaseUrl to your Fastvue Reporter API base URL.
# 2. Set $retentionUnit to "Days", "Months", or "Years".
# 3. Set $retentionValue to the number of units to keep.
# 4. Run once with $dryRun = $true to verify what would be deleted.
# 5. Set $dryRun = $false to perform deletion.
#
# Endpoints used:
# - Reports.GetReports
# - Reports.DeleteReports
#
# Requires the account running this script to have access to report APIs.

# -------------------- User-Defined Variables --------------------

# Fastvue Reporter API base URL.
$apiBaseUrl = "http://your_fastvue_site/_/api?f="

# Retention unit: "Days", "Months", or "Years"
$retentionUnit = "Months"

# Number of units to keep. Example: 3 with "Months" keeps the last 3 months.
$retentionValue = 3

# Safety mode: when $true, logs what would be deleted without deleting.
$dryRun = $true

# Number of report IDs to send per delete request.
$deleteBatchSize = 100

# -------------------- Script Begins --------------------

$scriptDirectory = $PSScriptRoot
$scriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)

function Get-AvailableLogPath {
    param($basePath)

    if (-not (Test-Path $basePath)) {
        return $basePath
    }

    try {
        $fileStream = [System.IO.File]::Open($basePath, 'Open', 'Write', 'None')
        $fileStream.Close()
        $fileStream.Dispose()
        return $basePath
    } catch {
        $counter = 1
        $directory = [System.IO.Path]::GetDirectoryName($basePath)
        $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($basePath)
        $extension = [System.IO.Path]::GetExtension($basePath)

        do {
            $newPath = Join-Path $directory "$baseFileName-$counter$extension"
            if (-not (Test-Path $newPath)) {
                return $newPath
            }

            try {
                $fileStream = [System.IO.File]::Open($newPath, 'Open', 'Write', 'None')
                $fileStream.Close()
                $fileStream.Dispose()
                return $newPath
            } catch {
                $counter++
            }
        } while ($true)
    }
}

$baseLogFile = Join-Path $scriptDirectory "$scriptName-output.log"
$script:logFile = Get-AvailableLogPath $baseLogFile

function Write-Log {
    param($Message)

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $Message"

    Write-Host $logMessage

    try {
        Add-Content -Path $script:logFile -Value $logMessage -ErrorAction Stop
    } catch {
        $script:logFile = Get-AvailableLogPath $baseLogFile
        try {
            Add-Content -Path $script:logFile -Value "=== Log continued from previous file ===" -ErrorAction Stop
            Add-Content -Path $script:logFile -Value $logMessage -ErrorAction Stop
            Write-Host "Note: Switched to new log file: $($script:logFile)"
        } catch {
            Write-Host "Warning: Unable to write to any log file: $_"
        }
    }
}

function Log-Error {
    param (
        [string]$message,
        [string]$errorDetail
    )
    Write-Log "ERROR: $message. Details: $errorDetail"
    exit 1
}

function Get-UserCredential {
    if ([Environment]::UserInteractive) {
        Write-Log "Please enter your credentials:"
        return Microsoft.PowerShell.Security\Get-Credential
    } else {
        Write-Log "Running in non-interactive mode. Using default credentials."
        return $null
    }
}

function Invoke-ReporterApi {
    param(
        [string]$endpoint,
        [ValidateSet("GET", "POST")]
        [string]$method = "POST",
        [string]$body = "{}"
    )

    $uri = "$apiBaseUrl$endpoint"
    try {
        if ($credentials) {
            if ($method -eq "POST") {
                return Invoke-RestMethod -Uri $uri -Method POST -Credential $credentials @httpAuthParams -ContentType "application/json" -Body $body
            } else {
                return Invoke-RestMethod -Uri $uri -Method GET -Credential $credentials @httpAuthParams
            }
        } else {
            if ($method -eq "POST") {
                return Invoke-RestMethod -Uri $uri -Method POST -UseDefaultCredentials @httpAuthParams -ContentType "application/json" -Body $body
            } else {
                return Invoke-RestMethod -Uri $uri -Method GET -UseDefaultCredentials @httpAuthParams
            }
        }
    } catch {
        Log-Error "API call failed for endpoint '$endpoint'" $_
    }
}

function Get-CutoffDate {
    param(
        [ValidateSet("Days", "Months", "Years")]
        [string]$Unit,
        [int]$Value
    )

    if ($Value -lt 1) {
        Log-Error "Invalid retention value" "retentionValue must be 1 or greater."
    }

    $now = Get-Date
    switch ($Unit) {
        "Days" { return $now.AddDays(-1 * $Value) }
        "Months" { return $now.AddMonths(-1 * $Value) }
        "Years" { return $now.AddYears(-1 * $Value) }
        default { Log-Error "Invalid retention unit" "Unit must be Days, Months, or Years." }
    }
}

function Convert-FromUnixTimeMilliseconds {
    param([object]$milliseconds)

    if ($null -eq $milliseconds) {
        return $null
    }

    try {
        $msValue = [int64]$milliseconds
        if ($msValue -le 0) {
            return $null
        }
        return [DateTimeOffset]::FromUnixTimeMilliseconds($msValue).LocalDateTime
    } catch {
        return $null
    }
}

function Resolve-ReportItems {
    param([object]$apiResponse)

    # Handles both possible shapes:
    # 1) { Status: 0, Data: [ { ID, ... }, ... ] }
    # 2) { Status: 0, Data: [ { Status: 0, Data: { ID, ... } }, ... ] }
    if ($null -eq $apiResponse -or $null -eq $apiResponse.Data) {
        return @()
    }

    $items = @()
    foreach ($entry in @($apiResponse.Data)) {
        if ($null -ne $entry.ID) {
            $items += $entry
        } elseif ($null -ne $entry.Data -and $null -ne $entry.Data.ID) {
            $items += $entry.Data
        }
    }

    return $items
}

# Check whether the API URL uses an unencrypted (http://) connection.
# $httpAuthParams is added to every API request below.
$httpAuthParams = @{}
if ($apiBaseUrl -match "^\s*http://") {
    Write-Log "Warning: The Fastvue API URL uses http://, so this connection is not encrypted. If this script runs on the Fastvue server itself, you can ignore this. Otherwise, consider using an https:// URL to protect your credentials and data in transit."

    # PowerShell 6 and later refuse to send credentials over http:// unless this is explicitly allowed.
    # Windows PowerShell 5.1 has no such check (and does not have this parameter).
    if ($PSVersionTable.PSVersion.Major -ge 6) {
        $httpAuthParams["AllowUnencryptedAuthentication"] = $true
    }
}

$credentials = Get-UserCredential

Write-Log "Starting report cleanup."
Write-Log "Retention: keep last $retentionValue $retentionUnit"
Write-Log "Dry run: $dryRun"

$cutoffDate = Get-CutoffDate -Unit $retentionUnit -Value $retentionValue
Write-Log "Cutoff date/time: $($cutoffDate.ToString('yyyy-MM-dd HH:mm:ss'))"

# Fastvue typically supports POST with an empty body for API functions.
$getReportsResponse = Invoke-ReporterApi -endpoint "Reports.GetReports" -method "POST" -body "{}"

if ($null -eq $getReportsResponse -or $getReportsResponse.Status -ne 0) {
    Log-Error "Failed to retrieve reports" ($getReportsResponse | ConvertTo-Json -Depth 5)
}

$reports = Resolve-ReportItems -apiResponse $getReportsResponse
Write-Log "Total reports returned: $($reports.Count)"

if ($reports.Count -eq 0) {
    Write-Log "No reports found. Nothing to do."
    exit 0
}

$candidateReports = @()
$missingDateCount = 0

foreach ($report in $reports) {
    $generatedDate = Convert-FromUnixTimeMilliseconds $report.GeneratedDate
    if ($null -eq $generatedDate) {
        $missingDateCount++
        continue
    }

    if ($generatedDate -lt $cutoffDate) {
        $candidateReports += [PSCustomObject]@{
            ID            = $report.ID
            Title         = $report.Title
            GeneratedDate = $generatedDate
        }
    }
}

Write-Log "Reports without valid GeneratedDate: $missingDateCount"
Write-Log "Reports older than cutoff: $($candidateReports.Count)"

if ($candidateReports.Count -eq 0) {
    Write-Log "No reports are older than cutoff. Nothing to delete."
    exit 0
}

foreach ($r in $candidateReports | Sort-Object GeneratedDate) {
    Write-Log "Candidate: ID=$($r.ID), GeneratedDate=$($r.GeneratedDate.ToString('yyyy-MM-dd HH:mm:ss')), Title=$($r.Title)"
}

if ($dryRun) {
    Write-Log "Dry run enabled. No reports were deleted."
    exit 0
}

$idsToDelete = @($candidateReports | ForEach-Object { $_.ID } | Where-Object { $_ })
$effectiveBatchSize = if ($deleteBatchSize -lt 1) { 100 } else { $deleteBatchSize }

$totalDeleted = 0
$failedBatches = 0

for ($i = 0; $i -lt $idsToDelete.Count; $i += $effectiveBatchSize) {
    $end = [Math]::Min($i + $effectiveBatchSize - 1, $idsToDelete.Count - 1)
    $batch = @($idsToDelete[$i..$end])

    $payload = @{
        ReportIDs = $batch
    } | ConvertTo-Json -Depth 3

    Write-Log "Deleting batch with $($batch.Count) report IDs..."
    $deleteResponse = Invoke-ReporterApi -endpoint "Reports.DeleteReports" -method "POST" -body $payload

    if ($null -eq $deleteResponse) {
        $failedBatches++
        Write-Log "ERROR: Empty response from Reports.DeleteReports."
        continue
    }

    if ($deleteResponse.Status -eq 0) {
        $totalDeleted += $batch.Count
        Write-Log "Batch deleted successfully."
    } else {
        $failedBatches++
        Write-Log "ERROR: Delete batch failed. Response: $($deleteResponse | ConvertTo-Json -Depth 5)"
    }
}

Write-Log "Cleanup complete."
Write-Log "Requested deletion count: $($idsToDelete.Count)"
Write-Log "Deleted count (by successful batches): $totalDeleted"
Write-Log "Failed batches: $failedBatches"
Write-Log "Log file: $script:logFile"
