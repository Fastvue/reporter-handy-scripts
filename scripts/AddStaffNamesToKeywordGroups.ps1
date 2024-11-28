# This script queries Active Directory for a list of staff names and adds them to a keyword group. 
# This is useful for schools that want to know when students are searching for staff online to find their social media profiles. 
# Before scheduling the script, first run it in an elvated PowerShell window to install the required AD module. 

 # -------------------- User-Defined Variables --------------------

# Fastvue Reporter API base URL.
$apiBaseUrl = "http://your_fastvue_site/_/api?f="

# The Active Directory (AD) server FQDN (Fully Qualified Domain Name) or IP address.
$adServer = "dc.example.local"

# LDAP query for filtering AD users.
$ldapQuery = "(objectClass=user)"

# Security Groups to import users from. Specify multiple groups as an array. e.g. @("Admin Staff", "Teaching Staff").
$securityGroups = @()  # If empty, will return all users matching the LDAP query.

# AD attribute to use for populating the keywords.
$adAttribute = "DisplayName"

# The name of the keyword group to create or update in Fastvue Reporter.
$keywordGroupName = "Staff Names"

 # -------------------- Script Begins --------------------

# Get the directory where the script is located
$scriptDirectory = $PSScriptRoot
$scriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)

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
$baseLogFile = Join-Path $scriptDirectory "$scriptName-output.log"
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

# Function to log errors and exit
function Log-Error {
    param (
        [string]$message,
        [string]$errorDetail
    )
    Write-Log "ERROR: $message. Details: $errorDetail"
    exit 1
}

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

# Replace the existing credential management section with:
$credentials = Get-UserCredential

# -------------------- Check and Install Active Directory Module --------------------
function Check-Install-ADModule {
    if (Get-Module -ListAvailable -Name ActiveDirectory) {
        Write-Log "Active Directory module is already installed."
        Import-Module ActiveDirectory
    } else {
        Write-Log "Active Directory module not found. Attempting to install..."
        try {
            if (Get-WindowsFeature -Name RSAT-AD-PowerShell -ErrorAction SilentlyContinue) {
                Install-WindowsFeature -Name "RSAT-AD-PowerShell"
            } elseif (Get-WindowsCapability -Name Rsat.ActiveDirectory* -Online) {
                Add-WindowsCapability -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 -Online
            } else {
                Log-Error "Failed to install Active Directory module" "Manual installation required"
            }
            Import-Module ActiveDirectory
            Write-Log "Active Directory module installed and imported successfully."
        } catch {
            Log-Error "Failed to import the Active Directory module" $_
        }
    }
}

# Call the function to check and install the AD module if necessary
Check-Install-ADModule

# -------------------- Retrieve Users from Security Groups --------------------
function Get-ADUsersFromGroups {
    param (
        [array]$securityGroups,
        [string]$adServer
    )

    $users = @()

    if ($securityGroups.Count -gt 0) {
        try {
            Write-Log "Querying users from security groups: $($securityGroups -join ', ')"
            
            foreach ($group in $securityGroups) {
                try {
                    Write-Log "Querying security group: $group"
                    # Get members of the security group
                    $groupMembers = Get-ADGroupMember -Identity $group -Server $adServer | Where-Object { $_.objectClass -eq 'user' }
                    if ($groupMembers) {
                        Write-Log "Found $($groupMembers.Count) members in security group '$group'"
                        $users += $groupMembers
                    } else {
                        Write-Log "No members found in security group '$group'"
                    }
                } catch {
                    Log-Error "Error querying security group '$group'" $_
                }
            }
        } catch {
            Log-Error "Error retrieving users from Active Directory groups" $_
        }
    } else {
        Write-Log "No security groups specified."
    }

    return $users | Sort-Object -Unique
}

# -------------------- Apply LDAP Query Using Get-ADUser --------------------
function FilterUsersByLDAP {
    param (
        [array]$users,
        [string]$ldapQuery,
        [string]$adServer
    )

    $filteredUsers = @()

    if ($users.Count -gt 0 -and $ldapQuery) {
        Write-Log "Applying LDAP query filter to group members using Get-ADUser."
        try {
            $distinguishedNames = $users | ForEach-Object { $_.DistinguishedName }

            # Apply the LDAP filter using Get-ADUser
            $filteredUsers = Get-ADUser -LDAPFilter $ldapQuery -Server $adServer -Properties DisplayName | Where-Object {
                $distinguishedNames -contains $_.DistinguishedName
            }

            # Output the filtered users for debugging
            Write-Log "Filtered Users after LDAP Query:"
            $filteredUsers | ForEach-Object { Write-Log "DisplayName: $($_.DisplayName), DistinguishedName: $($_.DistinguishedName)" }

        } catch {
            Log-Error "Error applying LDAP filter to the retrieved users" $_
        }
    } else {
        Write-Log "No users to filter or no LDAP query specified."
    }

    return $filteredUsers | Sort-Object -Unique
}

# -------------------- Main Logic --------------------
# Step 1: Check if security groups are specified, if not, directly query AD with the LDAP filter.
if ($securityGroups.Count -eq 0) {
    Write-Log "No security groups specified. Querying AD directly using LDAP filter."

    try {
        # Query AD directly using the LDAP filter when no security groups are specified
        $filteredUsers = Get-ADUser -LDAPFilter $ldapQuery -Server $adServer -Properties DisplayName

        # Output the filtered users for debugging
        Write-Log "Filtered Users after LDAP Query:"
        $filteredUsers | ForEach-Object { Write-Log "DisplayName: $($_.DisplayName), DistinguishedName: $($_.DistinguishedName)" }

    } catch {
        Log-Error "Error querying AD directly using LDAP filter" $_
    }

    $displayNames = $filteredUsers | ForEach-Object { $_.DisplayName }

} else {
    # Step 2: Get users from security groups if specified.
    $groupUsers = Get-ADUsersFromGroups -securityGroups $securityGroups -adServer $adServer

    # Step 3: Filter users based on LDAP query.
    $filteredUsers = FilterUsersByLDAP -users $groupUsers -ldapQuery $ldapQuery -adServer $adServer
    $displayNames = $filteredUsers | ForEach-Object { $_.DisplayName }

}

Write-Log "Found $($displayNames.Count) unique users after applying LDAP filter."

# -------------------- Fastvue Reporter Keyword Group Logic --------------------

$getKeywordGroupsUrl = "$apiBaseUrl`Settings.Keywords.GetGroups"
$createKeywordGroupUrl = "$apiBaseUrl`Settings.Keywords.AddGroup"
$updateKeywordsUrl = "$apiBaseUrl`Settings.Keywords.UpdateIncludes"

try {
    Write-Log "Checking if the keyword group '$keywordGroupName' already exists..."
    if ($credentials) {
        $groupsResponse = Invoke-RestMethod -Uri $getKeywordGroupsUrl -Method POST -Credential $credentials -ContentType "application/json" -Body "{}"
    } else {
        # Use default credentials if running under Task Scheduler
        $groupsResponse = Invoke-RestMethod -Uri $getKeywordGroupsUrl -Method POST -UseDefaultCredentials -ContentType "application/json" -Body "{}"
    }
    # Check if the group exists by name
    $existingGroup = $groupsResponse.Data | Where-Object { $_.Name -eq $keywordGroupName }

    if ($existingGroup) {
        Write-Log "Keyword group '$keywordGroupName' exists. Group ID: $($existingGroup.ID)"
        $keywordGroupId = $existingGroup.ID
    } else {
        Write-Log "Keyword group '$keywordGroupName' does not exist. Proceeding to create the group."
        try {
            $createGroupData = @{ "Name" = $keywordGroupName } | ConvertTo-Json
            if ($credentials) {
                $createGroupResponse = Invoke-RestMethod -Uri $createKeywordGroupUrl -Method POST -Credential $credentials -ContentType "application/json" -Body $createGroupData
            } else {
                $createGroupResponse = Invoke-RestMethod -Uri $createKeywordGroupUrl -Method POST -UseDefaultCredentials -ContentType "application/json" -Body $createGroupData
            }

            if ($createGroupResponse.Status -eq 0) {
                $keywordGroupId = $createGroupResponse.Data.ID
                Write-Log "Keyword group created successfully. Group ID: $keywordGroupId"
            } else {
                Log-Error "Unexpected error while creating keyword group" $createGroupResponse
            }
        } catch {
            Log-Error "Error creating keyword group" $_
        }
    }
} catch {
    Log-Error "Error retrieving keyword groups from Fastvue Reporter" $_
}

# -------------------- Update Keywords --------------------
$keywords = @($displayNames | ForEach-Object {
    @{
        "Keyword" = $_
        "WholeWord" = $false
    }
})

Write-Log "Constructed keywords for update:"
$keywords | ForEach-Object { Write-Log "Keyword: $($_.Keyword), WholeWord: $($_.WholeWord)" }

try {
    $updateKeywordsData = @{
        "GroupID" = $keywordGroupId
        "Keywords" = $keywords
    } | ConvertTo-Json -Depth 3

    
    if ($credentials) {
        $updateResponse = Invoke-RestMethod -Uri $updateKeywordsUrl -Method POST -Credential $credentials -ContentType "application/json" -Body $updateKeywordsData
    } else {
        $updateResponse = Invoke-RestMethod -Uri $updateKeywordsUrl -Method POST -UseDefaultCredentials -ContentType "application/json" -Body $updateKeywordsData
    }

    if ($updateResponse.Status -eq 0 -and $updateResponse.Data.Success) {
        Write-Log "Keywords added successfully to group '$keywordGroupName'."
    } else {
        Log-Error "Failed to add keywords to the group" $updateResponse
    }
} catch {
    Log-Error "Error updating keywords in Fastvue Reporter" $_
}
 
