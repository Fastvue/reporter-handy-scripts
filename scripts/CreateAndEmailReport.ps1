# This script generates a report for a specified time range, 
# with optional filters and sends an email with the private link to the report, 
# as well as PDF and CSV attachments.
# Here's a video on how to use it: https://www.loom.com/share/21cfc712542d434d803c0034f6accad3?sid=0937bcde-9ce0-48a0-9669-45a66601cef7

# UPDATE: 04/02/2025 - Added report duration options and more flexible filtering options to the user modifiable part of the script (not detailed in video above)

# Define API Base URL
$apiBaseUrl = "http://your_fastvue_site/_/api?f="

# Define Report Type: Can be "Internet Usage", "Safeguarding", "IT Network and Security", "All Usage" or "Activity"
$reportType = "Internet Usage"  # Change this to the desired report type above

# Define whether to include the current day in the report
$includeCurrentDay = $false  # Set to $true to include today in the report

# Define the report duration
$reportDuration = "Week"  # Options: "Day", "Week", "Month", or specify a custom number of days like "14" for two weeks.

# Define time variables (Only affects the time on the first and last day of the report. Leave both values as 00:00:00 for 'all day') 
$startTime = "00:00:00"  # Start time of the report - leave as 00:00:00 to start at midnight
$endTime = "00:00:00"    # End time of the report - leave as 00:00:00 to end at midnight

# Define report title
$baseReportTitle = "$reportType Report: Class ABC"  # E.g. Internet Usage Report: Class ABC

# Define Email Recipient(s)
$emailRecipient = "your.email@example.com"  # Separate multiple emails with a comma (,) or semi-colon (;)

# Define Filters
# To apply no filters, leave the $reportFilters array empty: @()

# Available operators:
#   Equal, NotEqual, StartsWith, NotStartsWith, EndsWith, NotEndsWith,
#   Contains, NotContains, ContainsWholeWord, NotContainsWholeWord,
#   InKeywordGroup, NotInKeywordGroup, InSubnet, NotInSubnet,
#   GreaterThan, GreaterOrEqual, LessThan, LessOrEqual

$reportFilters = @(
    # Uncomment and modify the following lines to add filters.
    # @{
    #     "Field" = "Security Group"
    #     "Operator" = "Equal"
    #     "Values" = @("Group One", "Group Two", "Group Three")
    # },
    # @{
    #     "Field" = "Origin Domain"
    #     "Operator" = "NotEqual"
    #     "Values" = @("microsoft.com", "windowsupdate.com", "adobe.com")
    # }
)

#### Do not modify below this line (unless you know what you're doing!) ####

# Convert user-friendly filters to the required format
$filters = @()
if ($reportFilters.Count -gt 0) {
    foreach ($userFilter in $reportFilters) {
        $filters += @{
            "Type" = "Value"
            "Semantic" = $userFilter.Field  
            "Operator" = $userFilter.Operator
            "Values" = $userFilter.Values
        }
    }
}

# Calculate start and end dates based on the report duration
switch ($reportDuration) {
    "Day" {
        if ($includeCurrentDay) {
            $startDate = Get-Date
            $endDate = Get-Date
        } else {
            $startDate = (Get-Date).AddDays(-1)  # Yesterday
            $endDate = (Get-Date).AddDays(-1)    # Yesterday
        }
    }
    "Week" {
        if ($includeCurrentDay) {
            $endDate = Get-Date
            $startDate = $endDate.AddDays(-6)  # -6 for 7 days total
        } else {
            $endDate = (Get-Date).AddDays(-1)  # Yesterday
            $startDate = $endDate.AddDays(-6)  # -6 for 7 days total
        }
    }
    "Month" {
        if ($includeCurrentDay) {
            $endDate = Get-Date
            $startDate = (Get-Date -Year $endDate.Year -Month $endDate.Month -Day 1).AddMonths(-1)
        } else {
            # For previous month, go from 1st to last day of previous month
            $endDate = (Get-Date -Day 1).AddDays(-1)  # Last day of previous month
            $startDate = $endDate.AddDays(-($endDate.Day - 1))  # First day of previous month
        }
    }
    default {
        if ($reportDuration -as [int]) {
            $endDate = if ($includeCurrentDay) { Get-Date } else { (Get-Date).AddDays(-1) }
            $startDate = $endDate.AddDays(-([int]$reportDuration - 1))
        } else {
            throw "Invalid report duration specified. Use 'Day', 'Week', 'Month', or a number of days."
        }
    }
}

# Construct start and end date-times
$startDateTime = "$($startDate.ToString('yyyy-MM-dd')) $startTime"
if ($endTime -eq "00:00:00") {
    # Add one day to end date when end time is midnight to include the full end date
    $endDateTime = "$($endDate.AddDays(1).ToString('yyyy-MM-dd')) $endTime"
} else {
    $endDateTime = "$($endDate.ToString('yyyy-MM-dd')) $endTime"
}

# Construct report title
$reportTitle = if ($startTime -eq "00:00:00" -and $endTime -eq "00:00:00") {
    if ($startDate.Date -eq $endDate.Date) {
        "$baseReportTitle ($($startDate.ToString('yyyy-MM-dd')))"
    } else {
        "$baseReportTitle ($($startDate.ToString('yyyy-MM-dd')) - $($endDate.ToString('yyyy-MM-dd')))"
    }
} else {
    "$baseReportTitle ($startDateTime - $endDateTime)"
}

# Get the directory where the script is located
$scriptDirectory = $PSScriptRoot

# Function to get an available log file path
function Get-AvailableLogPath {
    param($basePath)
    
    # Try the base path first
    if (-not (Test-Path $basePath)) {
        return $basePath
    }
    
    # Check if existing file is locked
    try {
        $fileStream = [System.IO.File]::Open($basePath, 'Open', 'Write', 'None')
        $fileStream.Close()
        $fileStream.Dispose()
        return $basePath
    } catch {
        # If file is locked, try numbered versions
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

# Initialize log file path
$baseLogFile = Join-Path $scriptDirectory "CreateAndEmailReport-output.log"
$script:logFile = Get-AvailableLogPath $baseLogFile

# Function to write to both console and log file
function Write-Log {
    param($Message)
    
    # Get current timestamp
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    # Create the log message
    $logMessage = "$timestamp - $Message"
    
    # Write to console
    Write-Host $logMessage
    
    # Append to log file
    try {
        Add-Content -Path $script:logFile -Value $logMessage -ErrorAction Stop
    } catch {
        # If writing fails, try to get a new log file
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

$reportCreateUrl = "$apiBaseUrl`Reports.Create"
$reportGetStatusUrl = "$apiBaseUrl`Reports.GetReport&id="
$reportShareUrl = "$apiBaseUrl`Reports.ShareViaEmail"

# Determine the layout based on the report type
switch ($reportType) {
    "Internet Usage" { $layout = "overview-internet" }
    "Safeguarding" {$layout = "overview-safeguarding" }
    "IT Network and Security" { $layout = "overview-it" }
    "All Usage" {$layout = "companyOverview"}
    "Activity" { $layout = "detailedInvestigation" }
    default { $layout = "overview-internet" }  # Default to Internet Usage if no valid report type is provided
}

# Define the report creation data
$reportData = @{
    "Layout" = $layout
    "Title" = $reportTitle
    "StartDate" = $startDateTime
    "EndDate" = $endDateTime
    "Filter" = $filters  # Use the user-defined filters
} | ConvertTo-Json -Depth 3

# Function to get credentials
function Get-UserCredential {
    if ([Environment]::UserInteractive) {
        Write-Log "Please enter your credentials:"
        return Microsoft.PowerShell.Security\Get-Credential
    } else {
        Write-Log "Running in non-interactive mode. Using default credentials."
        return $null
    }
}

# Get credentials
$credentials = Get-UserCredential

# Create the report using the API
try {
    if ($credentials) {
        $response = Invoke-RestMethod -Uri $reportCreateUrl -Method POST -ContentType "application/json" -Body $reportData -Credential $credentials
    } else {
        $response = Invoke-RestMethod -Uri $reportCreateUrl -Method POST -ContentType "application/json" -Body $reportData -UseDefaultCredentials
    }
    if (-not $response -or -not $response.Data) {
        throw "Failed to create the report. No response or invalid response received."
    }
    $reportId = $response.Data
    Write-Log "Report created successfully. Report ID: $reportId"
} catch {
    Write-Log "Error: Failed to create the report. $_"
    exit 1
}

# Function to check the report status
function Check-ReportStatus($reportId) {
    $getStatusUrl = "$reportGetStatusUrl$reportId"
    try {
        if ($credentials) {
            $statusResponse = Invoke-RestMethod -Uri $getStatusUrl -Method GET -Credential $credentials
        } else {
            $statusResponse = Invoke-RestMethod -Uri $getStatusUrl -Method GET -UseDefaultCredentials
        }
        
        if (-not $statusResponse -or -not $statusResponse.Data) {
            throw "Failed to retrieve report status. No response or invalid response received."
        }
        return $statusResponse
    } catch {
        Write-Log "Error: Failed to retrieve report status. $_"
        exit 1
    }
}

# Define polling interval (in seconds) for checking report status
$pollingInterval = 5  # In seconds. Adjust this interval if needed

# Polling the status of the report until it's completed
$status = "Processing"
while ($status -eq "Processing") {
    $reportStatus = Check-ReportStatus $reportId
    $processProgress = $reportStatus.Data.ProcessProgress * 100
    $status = $reportStatus.Data.ProcessStatus
    $statusMessage = $reportStatus.Data.ProcessStatusMessage
    Write-Log "Report Status: $statusMessage ($([math]::Round($processProgress, 2))%)"
    
    if ($status -eq "Completed") {
        Write-Log "Report generation completed successfully."
        break
    }
    elseif ($status -eq "Failed" -or $reportStatus.Data.ProcessError) {
        Write-Log "Report generation failed. Error: $($reportStatus.Data.ProcessError)"
        exit 1
    }

    # Wait for the defined polling interval before checking again
    Start-Sleep -Seconds $pollingInterval
}

# Define the email sharing data
$emailData = @{
    "ReportID" = $reportId
    "EmailTo" = $emailRecipient
    "UseLink" = $true
    "IsPrivate" = $true
    "AttachFormats" = @(
        @{
            "Format" = "PDF"
            "Options" = @{
                "IncludeActivityDetails" = $false
            }
        },
        @{
            "Format" = "CSV"
            "Options" = @{
                "IncludeActivityDetails" = $true
            }
        }
    )
} | ConvertTo-Json -Depth 3

# Share the report via email
try {
    if ($credentials) {
        $emailResponse = Invoke-RestMethod -Uri $reportShareUrl -Method POST -ContentType "application/json" -Body $emailData -Credential $credentials
    } else {
        $emailResponse = Invoke-RestMethod -Uri $reportShareUrl -Method POST -ContentType "application/json" -Body $emailData -UseDefaultCredentials
    }
    
    if (-not $emailResponse) {
        throw "Failed to email the report. No response received."
    }
    Write-Log "Report emailed successfully to: $emailRecipient"
} catch {
    Write-Log "Error: Failed to email the report. $_"
    exit 1
}
