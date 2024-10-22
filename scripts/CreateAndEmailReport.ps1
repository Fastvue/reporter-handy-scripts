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

# Define the log file path in the same directory as the script
$logFile = Join-Path $scriptDirectory "CreateAndEmailReport-output.log"

# Redirect output and errors to the log file
Start-Transcript -Path $logFile

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

# Function to check if credentials are needed
function Get-OptionalCredential {
    # If the script is running with credentials (such as from Task Scheduler), skip prompting.
    if (-not $credentials) {
        # Check if the script is running in an interactive session (e.g., manually by a user)
        if ($Host.Name -ne "ConsoleHost") {
            # Prompt for credentials if not provided and running interactively
            Write-Host "Credentials not detected. Prompting for credentials..."
            return Get-Credential
        } else {
            # Running under Task Scheduler or a non-interactive session. Do not prompt for credentials.
            Write-Host "No credentials provided, assuming script is running with the correct account context."
            return $null
        }
    }
}

# Check if the script is running under Windows Task Scheduler or being executed manually
$credentials = $null
$credentials = Get-OptionalCredential

# Create the report using the API
try {
    if ($credentials) {
        # Use credentials if available
        $response = Invoke-RestMethod -Uri $reportCreateUrl -Method POST -ContentType "application/json" -Body $reportData -Credential $credentials
    } else {
        # No credentials, rely on default Windows Task Scheduler account
        $response = Invoke-RestMethod -Uri $reportCreateUrl -Method POST -ContentType "application/json" -Body $reportData
    }

    if (-not $response -or -not $response.Data) {
        throw "Failed to create the report. No response or invalid response received."
    }
    $reportId = $response.Data
    Write-Host "Report created successfully. Report ID: $reportId"
} catch {
    Write-Host "Error: Failed to create the report. $_"
    Stop-Transcript
    exit 1
}

# Function to check the report status
function Check-ReportStatus($reportId) {
    $getStatusUrl = "$reportGetStatusUrl$reportId"
    try {
        if ($credentials) {
            # Use credentials if available
            $statusResponse = Invoke-RestMethod -Uri $getStatusUrl -Method GET -Credential $credentials
        } else {
            # No credentials, rely on default Windows Task Scheduler account
            $statusResponse = Invoke-RestMethod -Uri $getStatusUrl -Method GET
        }

        if (-not $statusResponse -or -not $statusResponse.Data) {
            throw "Failed to retrieve report status. No response or invalid response received."
        }
        return $statusResponse
    } catch {
        Write-Host "Error: Failed to retrieve report status. $_"
        Stop-Transcript
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
    Write-Host "Report Status: $statusMessage ($([math]::Round($processProgress, 2))%)"
    
    if ($status -eq "Completed") {
        Write-Host "Report generation completed successfully."
        break
    }
    elseif ($status -eq "Failed" -or $reportStatus.Data.ProcessError) {
        Write-Host "Report generation failed. Error: $($reportStatus.Data.ProcessError)"
        Stop-Transcript
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
        # Use credentials if available
        $emailResponse = Invoke-RestMethod -Uri $reportShareUrl -Method POST -ContentType "application/json" -Body $emailData -Credential $credentials
    } else {
        # No credentials, rely on default Windows Task Scheduler account
        $emailResponse = Invoke-RestMethod -Uri $reportShareUrl -Method POST -ContentType "application/json" -Body $emailData
    }

    if (-not $emailResponse) {
        throw "Failed to email the report. No response received."
    }
    Write-Host "Report emailed successfully to: $emailRecipient"
} catch {
    Write-Host "Error: Failed to email the report. $_"
    Stop-Transcript
    exit 1
}

# End the transcript to stop logging
Stop-Transcript
 
