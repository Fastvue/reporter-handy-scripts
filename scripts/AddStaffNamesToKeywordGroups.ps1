 # -------------------- User-Defined Variables --------------------

# Fastvue Reporter API base URL.
$apiBaseUrl = "http://192.168.20.40/fortigate/_/api?f="

# The Active Directory (AD) server FQDN (Fully Qualified Domain Name) or IP address.
$adServer = "dc.exmaple.local"

# LDAP Filter for querying AD users.
$ldapQuery = "(objectClass=user)"

# Security Groups to import users from. Specify multiple groups as an array. e.g. @("Admin Staff", "Teaching Staff").
$securityGroups = @()  # If empty, will use only LDAP filter

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
    } catch {
        Write-Host "Transcription not supported or failed. Proceeding without transcription."
    }
}

# Call to start logging
Start-Logging

# -------------------- Check and Install Active Directory Module --------------------
function Check-Install-ADModule {
    if (Get-Module -ListAvailable -Name ActiveDirectory) {
        Import-Module ActiveDirectory
    } else {
        try {
            if (Get-WindowsFeature -Name RSAT-AD-PowerShell -ErrorAction SilentlyContinue) {
                Install-WindowsFeature -Name "RSAT-AD-PowerShell"
            } elseif (Get-WindowsCapability -Name Rsat.ActiveDirectory* -Online) {
                Add-WindowsCapability -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 -Online
            } else {
                Log-Error "Failed to install Active Directory module" "Manual installation required"
            }
            Import-Module ActiveDirectory
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
        foreach ($group in $securityGroups) {
            try {
                # Get members of the security group
                $groupMembers = Get-ADGroupMember -Identity $group -Server $adServer | Where-Object { $_.objectClass -eq 'user' }
                $users += $groupMembers
            } catch {
                Log-Error "Error querying security group '$group'" $_
            }
        }
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
        try {
            $distinguishedNames = $users | ForEach-Object { $_.DistinguishedName }
            $filteredUsers = Get-ADUser -LDAPFilter $ldapQuery -Server $adServer -Properties DisplayName | Where-Object {
                $distinguishedNames -contains $_.DistinguishedName
            }
        } catch {
            Log-Error "Error applying LDAP filter to the retrieved users" $_
        }
    }

    return $filteredUsers | Sort-Object -Unique
}

# -------------------- Main Logic --------------------
# Step 1: If no security groups specified, query AD directly using LDAP filter.
if ($securityGroups.Count -eq 0) {
    try {
        $filteredUsers = Get-ADUser -LDAPFilter $ldapQuery -Server $adServer -Properties DisplayName
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

# -------------------- Fastvue Reporter Keyword Group Logic --------------------

$getKeywordGroupsUrl = "$apiBaseUrl`Settings.Keywords.GetGroups"
$createKeywordGroupUrl = "$apiBaseUrl`Settings.Keywords.AddGroup"
$updateKeywordsUrl = "$apiBaseUrl`Settings.Keywords.UpdateIncludes"

try {
    $groupsResponse = Invoke-RestMethod -Uri $getKeywordGroupsUrl -Method POST -ContentType "application/json" -Body "{}"
    $existingGroup = $groupsResponse.Data | Where-Object { $_.Name -eq $keywordGroupName }

    if ($existingGroup) {
        $keywordGroupId = $existingGroup.ID
    } else {
        try {
            $createGroupData = @{ "Name" = $keywordGroupName } | ConvertTo-Json
            $createGroupResponse = Invoke-RestMethod -Uri $createKeywordGroupUrl -Method POST -ContentType "application/json" -Body $createGroupData

            if ($createGroupResponse.Status -eq 0) {
                $keywordGroupId = $createGroupResponse.Data.ID
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

try {
    $updateKeywordsData = @{
        "GroupID" = $keywordGroupId
        "Keywords" = $keywords  # Ensure it's always an array
    } | ConvertTo-Json -Depth 3

    $updateResponse = Invoke-RestMethod -Uri $updateKeywordsUrl -Method POST -ContentType "application/json" -Body $updateKeywordsData

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
 
