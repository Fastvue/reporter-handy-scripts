# This script generates a report for a specified time range in the day, 
# filtered by a list of users and sends an email with the private link to the report, 
# as well as PDF and CSV attachments.
# Here's a video on how to use it: https://www.loom.com/share/21cfc712542d434d803c0034f6accad3?sid=0937bcde-9ce0-48a0-9669-45a66601cef7

# Define API Base URL
$apiBaseUrl = "http://your_fastvue_site/_/api?f="

# Define Report Type: Can be "Internet Usage", "Safeguarding", "IT Network and Security", "All Usage" or "Activity"
$reportType = "Internet Usage"  # Change this to the desired report type above

# Define time variables
$startTime = "11:00:00"  # Start time of the report
$endTime = "12:00:00"    # End time of the report

# Define report title
$baseReportTitle = "$reportType Report: Class ABC"  # E.g. Internet Usage Report: Class ABC

# Define Email Recipient(s)
$emailRecipient = "your.email@example.com"  # Separate multiple emails with a comma (,) or semi-colon (;)

# Define Filter Values (List of Users)
$filterUsers = @("User One", "User Two", "User Three")  # Enter users exactly as you see them in Fastvue Reporter.

#### Do not modify below this line (unless you know what you're doing!) ####

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

# Get today's date in the desired format
$today = Get-Date -Format "yyyy-MM-dd"

# Construct start and end date-times
$startDateTime = "$today $startTime"
$endDateTime = "$today $endTime"

# Determine the layout based on the report type
switch ($reportType) {
    "Internet Usage" { $layout = "overview-internet" }
    "Safeguarding" {$layout = "overview-safeguarding" }
    "IT Network and Security" { $layout = "overview-it" }
    "All Usage" {$layout = "companyOverview"}
    "Activity" { $layout = "detailedInvestigation" }
    default { $layout = "overview-internet" }  # Default to Internet Usage if no valid report type is provided
}

$reportTitle = "$baseReportTitle ($startDateTime - $endDateTime)" # Append date/times to report title

# Define the report creation data
$reportData = @{
    "Layout" = $layout
    "Title" = $reportTitle
    "StartDate" = $startDateTime
    "EndDate" = $endDateTime
    "Filter" = @(
        @{
            "Type" = "Value"
            "Semantic" = "User"
            "Operator" = "Equal"
            "Values" = $filterUsers  # Use the user-defined filter values
        }
    )
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
