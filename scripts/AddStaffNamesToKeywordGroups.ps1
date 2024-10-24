# This script queries Active Directory for a list of staff names and adds them to a keyword group. 
# This is useful for schools that want to know when students are searching for staff online to find their social media profiles. 
# Before scheduling the script, first run it in an elvated PowerShell window to install the required AD module. 

 # -------------------- User-Defined Variables --------------------

# Fastvue Reporter API base URL.
$apiBaseUrl = "http://your_fastvue_site/_/api?f="

# The Active Directory (AD) server FQDN (Fully Qualified Domain Name) or IP address.
$adServer = "dc.exmaple.local"

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
$logFile = Join-Path $scriptDirectory "$scriptName-output.log"
$transcriptionRunning = $false
$useCredential = $false

# -------------------- Centralized Error Logging --------------------
function Log-Error {
    param (
        [string]$message,
        [string]$errorDetail
    )
    Write-Host "ERROR: $message. Details: $errorDetail" -ForegroundColor Red
    if ($transcriptionRunning) { Stop-Transcript }
    exit 1
}

# -------------------- Start Logging --------------------
function Start-Logging {
    try {
        Start-Transcript -Path $logFile
        $transcriptionRunning = $true
        Write-Host "Transcription started, logging to $logFile"
    } catch {
        Write-Host "Transcription not supported or failed. Proceeding without transcription."
    }
}

# Call to start logging
Start-Logging

# -------------------- Check and Install Active Directory Module --------------------
function Check-Install-ADModule {
    if (Get-Module -ListAvailable -Name ActiveDirectory) {
        Write-Host "Active Directory module is already installed."
        Import-Module ActiveDirectory
    } else {
        Write-Host "Active Directory module not found. Attempting to install..."
        try {
            if (Get-WindowsFeature -Name RSAT-AD-PowerShell -ErrorAction SilentlyContinue) {
                Install-WindowsFeature -Name "RSAT-AD-PowerShell"
            } elseif (Get-WindowsCapability -Name Rsat.ActiveDirectory* -Online) {
                Add-WindowsCapability -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 -Online
            } else {
                Log-Error "Failed to install Active Directory module" "Manual installation required"
            }
            Import-Module ActiveDirectory
            Write-Host "Active Directory module installed and imported successfully."
        } catch {
            Log-Error "Failed to import the Active Directory module" $_
        }
    }
}

# Call the function to check and install the AD module if necessary
Check-Install-ADModule

# -------------------- Credential Management --------------------
# Function to determine if the script is running interactively
function Is-Interactive {
    return ($Host.UI.RawUI.KeyAvailable -eq $true)
}

# Function to get credentials only if running interactively
function Get-Credentials {
    if (Is-Interactive) {
        Write-Host "Credentials not detected. Prompting for credentials..."
        $credential = Get-Credential
        Write-Host "Username entered: $($credential.UserName)"
        $useCredential = $true
        return $credential
    } else {
        Write-Host "Running in non-interactive mode. Using default credentials."
        return $null
    }
}

# Get credentials (only prompts if running interactively)
$credential = Get-Credentials

# -------------------- Retrieve Users from Security Groups --------------------
function Get-ADUsersFromGroups {
    param (
        [array]$securityGroups,
        [string]$adServer
    )

    $users = @()

    if ($securityGroups.Count -gt 0) {
        try {
            Write-Host "Querying users from security groups: $($securityGroups -join ', ')"
            
            foreach ($group in $securityGroups) {
                try {
                    Write-Host "Querying security group: $group"
                    # Get members of the security group
                    $groupMembers = Get-ADGroupMember -Identity $group -Server $adServer | Where-Object { $_.objectClass -eq 'user' }
                    if ($groupMembers) {
                        Write-Host "Found $($groupMembers.Count) members in security group '$group'"
                        $users += $groupMembers
                    } else {
                        Write-Host "No members found in security group '$group'"
                    }
                } catch {
                    Log-Error "Error querying security group '$group'" $_
                }
            }
        } catch {
            Log-Error "Error retrieving users from Active Directory groups" $_
        }
    } else {
        Write-Host "No security groups specified."
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
        Write-Host "Applying LDAP query filter to group members using Get-ADUser."
        try {
            $distinguishedNames = $users | ForEach-Object { $_.DistinguishedName }

            # Apply the LDAP filter using Get-ADUser
            $filteredUsers = Get-ADUser -LDAPFilter $ldapQuery -Server $adServer -Properties DisplayName | Where-Object {
                $distinguishedNames -contains $_.DistinguishedName
            }

            # Output the filtered users for debugging
            Write-Host "Filtered Users after LDAP Query:"
            $filteredUsers | ForEach-Object { Write-Host "DisplayName: $($_.DisplayName), DistinguishedName: $($_.DistinguishedName)" }

        } catch {
            Log-Error "Error applying LDAP filter to the retrieved users" $_
        }
    } else {
        Write-Host "No users to filter or no LDAP query specified."
    }

    return $filteredUsers | Sort-Object -Unique
}

# -------------------- Main Logic --------------------
# Step 1: Check if security groups are specified, if not, directly query AD with the LDAP filter.
if ($securityGroups.Count -eq 0) {
    Write-Host "No security groups specified. Querying AD directly using LDAP filter."

    try {
        # Query AD directly using the LDAP filter when no security groups are specified
        $filteredUsers = Get-ADUser -LDAPFilter $ldapQuery -Server $adServer -Properties DisplayName

        # Output the filtered users for debugging
        Write-Host "Filtered Users after LDAP Query:"
        $filteredUsers | ForEach-Object { Write-Host "DisplayName: $($_.DisplayName), DistinguishedName: $($_.DistinguishedName)" }

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

Write-Host "Found $($displayNames.Count) unique users after applying LDAP filter."

# -------------------- Fastvue Reporter Keyword Group Logic --------------------

$getKeywordGroupsUrl = "$apiBaseUrl`Settings.Keywords.GetGroups"
$createKeywordGroupUrl = "$apiBaseUrl`Settings.Keywords.AddGroup"
$updateKeywordsUrl = "$apiBaseUrl`Settings.Keywords.UpdateIncludes"

try {
    Write-Host "Checking if the keyword group '$keywordGroupName' already exists..."
    if ($useCredential -and $credential) {
        $groupsResponse = Invoke-RestMethod -Uri $getKeywordGroupsUrl -Method POST -Credential $credential -ContentType "application/json" -Body "{}"
    } else {
        # Use default credentials if running under Task Scheduler
        $groupsResponse = Invoke-RestMethod -Uri $getKeywordGroupsUrl -Method POST -UseDefaultCredentials -ContentType "application/json" -Body "{}"
    }
    # Check if the group exists by name
    $existingGroup = $groupsResponse.Data | Where-Object { $_.Name -eq $keywordGroupName }

    if ($existingGroup) {
        Write-Host "Keyword group '$keywordGroupName' exists. Group ID: $($existingGroup.ID)"
        $keywordGroupId = $existingGroup.ID
    } else {
        Write-Host "Keyword group '$keywordGroupName' does not exist. Proceeding to create the group."
        try {
            $createGroupData = @{ "Name" = $keywordGroupName } | ConvertTo-Json
            if ($useCredential) {
                $createGroupResponse = Invoke-RestMethod -Uri $createKeywordGroupUrl -Method POST -Credential $credential -ContentType "application/json" -Body $createGroupData
            } else {
                $createGroupResponse = Invoke-RestMethod -Uri $createKeywordGroupUrl -Method POST -UseDefaultCredentials -ContentType "application/json" -Body $createGroupData
            }

            if ($createGroupResponse.Status -eq 0) {
                $keywordGroupId = $createGroupResponse.Data.ID
                Write-Host "Keyword group created successfully. Group ID: $keywordGroupId"
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

Write-Host "Constructed keywords for update:"
$keywords | ForEach-Object { Write-Host "Keyword: $($_.Keyword), WholeWord: $($_.WholeWord)" }

try {
    $updateKeywordsData = @{
        "GroupID" = $keywordGroupId
        "Keywords" = $keywords
    } | ConvertTo-Json -Depth 3

    
    if ($useCredential) {
        $updateResponse = Invoke-RestMethod -Uri $updateKeywordsUrl -Method POST -Credential $credential -ContentType "application/json" -Body $updateKeywordsData
    } else {
        $updateResponse = Invoke-RestMethod -Uri $updateKeywordsUrl -Method POST -UseDefaultCredentials -ContentType "application/json" -Body $updateKeywordsData
    }

    if ($updateResponse.Status -eq 0 -and $updateResponse.Data.Success) {
        Write-Host "Keywords added successfully to group '$keywordGroupName'."
    } else {
        Log-Error "Failed to add keywords to the group" $updateResponse
    }
} catch {
    Log-Error "Error updating keywords in Fastvue Reporter" $_
}

# -------------------- Stop Logging --------------------
if ($transcriptionRunning) {
    Stop-Transcript
}
 
