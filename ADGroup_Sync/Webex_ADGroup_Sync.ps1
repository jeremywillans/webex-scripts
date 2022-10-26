################################
#
# AD Group Sync for Webex
# Written by Jeremy Willans
# https://github.com/jeremywillans/webex-scripts
# Version: 1.2
#
# USE AT OWN RISK, SCRIPT NOT FULLY TESTED NOR SUPPLIED WITH ANY GURANTEE
#
# Usage - Allows you to keep AD Group and Webex Spaces Memberships in sync
# Example - powershell -executionpolicy bypass path\to\Webex_ADGroup_Sync.ps1 <ARGUMENTS>
#
# Parameters:
# -Debug (default false) - Detailed information during process
#
# Script Variables
# WebexAuth (required) - Webex Authentication token for Bot
# ReportId (required) - User/Space ID to send error status updates
# ExemptUsers (optional) - Array of users to be excluded from the Sync removal process
# LogDir (default $PSScriptRoot) - Appending Transcript Log of process
# CSV (default groups.csv) - Source of AD to Webex Space Mapping (using RoomId)
# DaysBack (default -8) - Numbers of days to keep audit logs for
#
# Change History
# 1.0 20210318 Initial Release
# 1.1 20210629 Adds Support for FollowRelLink (Large Spaces) and convert to Markdown
# 1.2 20221026 Improve AD Error Handling, explicitly include Mail attribute and formatting fixes
#
# NOTE: This requires Powershell 7 to function
#
###
Param (
    [Switch]$Debug
)
###

If ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Output "This script requires Powershell Version 7+"
    Write-Output ""
    Exit 1
}


$WebexAuth = " << BOT TOKEN >> "
$ReportId = " << DEBUG SPACE/PERSON ID >> "
$ExemptUsers = @('') # @('user1@example.com','user2@example.com','user3@example.com')
$LogDir = "$PSScriptRoot\Logs"
$CSV = "$PSScriptRoot\groups.csv"
$DaysBack = "-8"

# Update Debug Logging
$DebugPreference = "SilentlyContinue"
If ($Debug) {
    $DebugPreference = "Continue"
}

# Setup Transcript
$ErrorActionPreference = "SilentlyContinue"
Stop-Transcript | Out-Null
$ErrorActionPreference = "Continue"
$Date = Get-Date -Format yyyy-MM-dd
$LogPath = "$LogDir\aadsynrule_auditlog-$date.log"
Start-Transcript -Path $LogPath -Append

# Prepare Webex Headers
$Authorization = 'Bearer ' + $WebexAuth
$Headers = @{'Content-Type' = 'application/json'; 'Authorization' = $Authorization }

# Required for Powershell REST Commands
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

Try {
    # Import Group List
    $GroupList = Import-Csv $CSV
    $GroupCount = $GroupList | Measure-Object
}
Catch {
    Write-Output "ERR: File not found: $CSV, aborting..."
    Exit 1
}

# Check Group Count
If ($GroupCount.Count -lt 1) {
    Write-Output "ERR: No spaces found in $CSV, aborting..."
    Exit 1
}
Write-Debug ("Groups Identified for AD Sync: " + $GroupCount.Count)

Try {
    # Check Bot is valid
    $ADSyncBot = Invoke-RestMethod -Headers $Headers -Uri https://webexapis.com/v1/people/me
    Write-Debug ($ADSyncBot.displayName + "bot ok!")
}
Catch {
    Write-Output "ERR: Bot Invalid, aborting..."
    Write-Debug $_.Exception.Message
    Write-Output ""
    Exit 1
}

# Prepare Error Tables
$ErrorResult = @()

Write-Output "Commencing AD Sync..."
# Commence Sync
ForEach ($Item in $GroupList) {

    # Trim the Variables
    $ADGroup = $Item.AD_Group.Trim()
    $WebexSpace = $Item.Webex_Space.Trim()

    $WebexUserList = @()
    $ADUserList = @()

    Try {
        # Check Space Moderation
        Write-Debug "[$ADGroup] Check Webex space moderated status"
        #$Body = @{"roomId" = "$WebexSpace";"personEmail" = $UserMail} | ConvertTo-Json
        $Result = Invoke-RestMethod -Headers $Headers -Uri https://webexapis.com/v1/rooms/$WebexSpace

        # Verify is a group space
        If ($Result.type -eq 'direct') {
            Write-Output "[$ADGroup] Not a Webex group space, aborting..."
            $ErrorResult += "[$ADGroup] Not a Webex group space"
            Continue
        }
        
        If ($Result.isLocked) {
            Write-Debug "[$ADGroup] Webex Space is moderated, checking if bot is moderator"
            
            # Get Membership Id
            $ADSyncBotId = $ADSyncBot.id
            $Result = Invoke-RestMethod -Headers $Headers -Uri "https://webexapis.com/v1/memberships?roomId=$WebexSpace&personId=$ADSyncBotId"
            $MembershipId = $Result.items[0].id

            # Verify Moderator Status from Membership
            $Result = Invoke-RestMethod -Headers $Headers -Uri "https://webexapis.com/v1/memberships/$MembershipId"
            
            If (!$Result.isModerator) {
                Write-Output "[$ADGroup] Bot is not moderator, aborting..."
                $ErrorResult += "[$ADGroup] Bot is not moderator"
                Continue
            }
            Write-Debug "[$ADGroup] Bot is moderator, continuing..."


        }
        Else {
            Write-Debug "[$ADGroup] Webex Space is not moderated"
        }

    }
    Catch {
        # Catch Error
        Write-Debug "[$ADGroup] Unable to get space status, is the bot a member?"
        Write-Debug ("[$ADGroup] Error Message: " + $_.Exception.Message)
        $ErrorResult += "[$ADGroup] Unable to get space status, is the bot a member?"
        Continue
                
    }

    # Get Webex Users
    $Result = Invoke-RestMethod -Headers $Headers -FollowRelLink -Uri https://webexapis.com/v1/memberships?max=1000"&"roomId=$WebexSpace 

    # Add Users to Array
    ForEach ($Item in $Result.items) {

        # Create List of Webex Users, excluding Bots
        If ($Item.personEmail -notmatch 'webex.bot') {
            $WebexUserList += $Item.personEmail
        }
    }

    $WebexCount = $WebexUserList.Count
    Write-Output "[$ADGroup] $WebexCount members currently in Webex"

    Try {
        # Get AD Group Members
        $ADUsers = Get-ADGroupMember -Identity $ADGroup -Recursive -ErrorAction Stop | ForEach-Object { Get-ADUser -Identity $_.DistinguishedName -Properties 'Mail' -ErrorAction Stop }
    }
    Catch {
        Write-Output "[$ADGroup] Unable to get group status"
        Write-Output ("[$ADGroup] " + $_.Exception.Message)
        Write-Debug ("[$ADGroup] " + $_.Exception)
        $ErrorResult += "[$ADGroup] Unable to get group status"
        Continue
    }
    
    # Recursive Check Users and Cross-reference with Webex Members
    ForEach ($User in $ADUsers) {

        # Check for User Email
        If (!$User.Mail) {
            Write-Debug "[$ADGroup] $User does not have an Email address"
            $ErrorResult += "[$ADGroup] $User does not have an Email address"
            Continue
        }

        # Select User Mail Attribute
        $UserMail = $User.Mail
        
        # Create List of AD Users
        $ADUserList += $UserMail

        # Check if user not in Webex List
        If ($WebexUserList -notcontains $UserMail) {
            Write-Debug "[$ADGroup] $UserMail Not on Webex List"

            Try {
                # Attempt Add User to Space
                Write-Debug "[$ADGroup] $UserMail attempting add to Webex Space"
                $Body = @{"roomId" = "$WebexSpace"; "personEmail" = $UserMail } | ConvertTo-Json
                $Result = Invoke-RestMethod -Method Post -Headers $Headers -Uri https://webexapis.com/v1/memberships -Body $Body

                $WebexUserList += $UserMail
                Write-Debug "[$ADGroup] $UserMail added to space"

            }
            Catch {
                # Catch Error
                Write-Debug "[$ADGroup] $User could not be added to space"
                Write-Debug ("[$ADGroup] $User Error Message: " + $_.Exception.Message)
                $ErrorResult += "[$ADGroup] $User could not be added to space, check email address and user exists in Webex"
                Continue
                
            }
        }
    }

    $ADCount = $ADUserList.Count
    Write-Output "[$ADGroup] $ADCount members currently in AD Group"

    ForEach ($User in $WebexUserList) {

        If ($ADUserList -notcontains $User) {
            Write-Debug "[$ADGroup] $User not in AD Group, removing..."

            # Check against exemption list
            If ($ExemptUsers -contains $User) { 
                Write-Debug "[$ADGroup] $User matched exempt list, ignoring..."
                Continue
            }

            Try {
                # Find Membership Id
                $Result = Invoke-RestMethod -Headers $Headers -Uri "https://webexapis.com/v1/memberships?roomId=$WebexSpace&personEmail=$User"
                $MembershipId = $Result.items[0].id

                # Remove Membership
                $Result = Invoke-RestMethod -Method Delete -Headers $Headers -Uri "https://webexapis.com/v1/memberships/$MembershipId"
                Write-Debug "[$ADGroup] $User removed from space."

            }
            Catch {
                # Catch Error
                Write-Debug "[$ADGroup] $User could not be removed from space"
                Write-Debug ("[$ADGroup] $User Error Message: " + $_.Exception.Message)
                $ErrorResult += "[$ADGroup] $User could not be removed from space"
                Continue
            }
        }
    }
}

# Error Reporting
If ($ErrorResult.Count -ne 0) {

    # Output Errors
    Write-Host 'Errors Encountered:'
    $ErrorResult | Out-String | Write-Host

    # Verify if ReportId was specified
    If (!$ReportId) {
        Write-Output 'ERR: No ReportId Specified'
    }
    Else {
        # Format Message
        $Html = ("<strong>AD Group Sync Report</strong><blockquote class=danger>" + ($ErrorResult -join '\n') + "</blockquote>")

        Try {
            # Test if ReportId is a spaceId
            Write-Debug "Testing ReportId as Space"
            $Result = Invoke-RestMethod -Headers $Headers -Uri https://webexapis.com/v1/memberships?roomId=$ReportId
            # If didnt error, post message
            Write-Debug "Posting message to Space"
            $Body = @{"roomId" = $ReportId; "html" = $Html } | ConvertTo-Json | ForEach-Object { [System.Text.RegularExpressions.Regex]::Unescape($_) }
            $Result = Invoke-RestMethod -Method Post -Headers $Headers -Uri https://webexapis.com/v1/messages -Body $Body
        }
        Catch {
            # Test if ReportId is a personId
            Write-Debug "Testing ReportId as PersonId"
            Try {

                $Result = Invoke-RestMethod -Headers $Headers -Uri https://webexapis.com/v1/people/$ReportId
                Write-Debug "Posting direct message"
                $Body = @{"toPersonId" =  $ReportId; "html" = $Html } | ConvertTo-Json | ForEach-Object { [System.Text.RegularExpressions.Regex]::Unescape($_) }
                $Result = Invoke-RestMethod -Method Post -Headers $Headers -Uri https://webexapis.com/v1/messages -Body $Body
            }
            Catch {
                Write-Output "Unable to send output report"
                Write-Debug ("Error Message: " + $_.Exception.Message)
            }
        }
    } 
}

# Prepare Old Logs Cleanup
$CurrentDate = Get-Date
$DatetoDelete = $CurrentDate.AddDays($DaysBack)

# Cleanup Script Logs
Write-Debug "Begin Script Log Cleanup"
Get-ChildItem $LogDir -Recurse | Where-Object { $_.LastWriteTime -lt $DatetoDelete } |
ForEach-Object {
    Try {
        Remove-Item $_.FullName -Recurse -ErrorAction Stop
        Write-Debug "Deleting file: $_"
    }
    Catch {
        Write-Output "Error deleting file: $_"
    }
}
Write-Debug "End Script Log Cleanup"

# Stop Transcript
Stop-Transcript | Out-Null