###############################
#
# AD SIP Proxy Address Updater Script
# Written by Jeremy Willans, Cisco Systems Australia.
# Date: 26-11-2020
#
# USE AT OWN RISK, SCRIPT NOT FULLY TESTED NOR SUPPLIED WITH ANY GURANTEE
#
# Parameters:
#
# -Force (default false) - Changes written to Active Directory
# -Debug (default false) - Detailed information during process
# -Credential (default false) - Allows different credentials to be specified at runtime
# -JobTitle (default false) - Requires Job title to be specificed for account to be updated
# -Group (default Domain Users) - AD Group used to select which users to update
# -Log (default sipproxy_auditlog.txt) - Appending Transcript Log of process
# -SkipTest (default false) - Allows skipping the prelim test
#
# Usage - Allows you to populate AD ProxyAddress Field with the required SIP:<Email Address> for proper Presence Integration with Webex
# Example - powershell -executionpolicy bypass path\to\Webex_Add_SIP_Proxy_Address.ps1 <ARGUMENTS>
#
###
Param (
    [Switch]$Force,
    [Switch]$Debug,
	[Switch]$Cedential,
    [Switch]$JobTitle,
    [Switch]$SkipTest,
	[String]$Group = "Domain Users",
	[String]$Log = "$PSScriptRoot\sipproxy_auditlog.txt"
 )
###

# Update Debug Logging
$DebugPreference = "SilentlyContinue"
If ($Debug) {
    $DebugPreference = "Continue"
}
 
# Results File
$Date = Get-Date -Format yyyy-MM-dd-hhmmss
$ResultsFile = "$PSScriptRoot\SIPProxyResults-$date.csv"

# Setup Transcript
$ErrorActionPreference="SilentlyContinue"
Stop-Transcript | Out-Null
$ErrorActionPreference = "Continue"
Start-Transcript -path $LogFile -append

# Clear Screen and present script parameters
cls
Write-Host
Write-Host "SIP Proxy Update Script"
Write-Host

# Request proposed credentials, if required
If ($Credential) {
    $Cred = Get-Credential -Message 'Please type your administritive credentials' -UserName (WhoAmI)
    If ($Cred -eq $null) {
        Write-Host
        Write-Host "No Credentials specified, aborting."
        Write-Host
        Break
    }
}

# Test Update AD Group Object
If (!$SkipTest) {
    Try 
    {
        Write-Debug "Test writing Description field back to Domain Users group..."
        If (!$Credential)
        {
            Get-ADGroup -Identity "Domain Users" -Properties Description | ForEach { Set-ADGroup -Identity "Domain Users" -Description $_.Description }
        
        } Else {
            Get-ADGroup -Credential $Cred -Identity "Domain Users" -Properties Description | ForEach { Set-ADGroup -Identity "Domain Users" -Description $_.Description }

        }
        Write-Debug "Test Successful."
    }
    Catch
    { 
        Write-Host "ERROR - Unable to update objects in AD, please check or run using -Credential to specify alternate credentials, aborting."
        Write-Host 
        Break
    }
}

# Get AD Info
$domainName = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Name

# Import User List
Try {

    # Import User List
    $userList = Get-ADGroupMember -Identity $Group
}
Catch
{

    Write-Debug "$Error[0]"
    Write-Host "AD Group: '$($Group)' does not exist, only the name 'Domain Users' etc needs to be specified with -Group <name>"
    Write-Host
    Break
}

# Check User Count before continuing
If ($userList.Count -eq 0) {
    Write-Host "AD Group: '$($Group)' does not contan any users. Please verify."
    Write-Host
    Break
}

# Output options
Write-Host "Audit Log: $LogFile"
Write-Host "Domain: $domainName"
Write-Host "AD Group: $Group"
Write-Host "Job Title Needed: $JobTitle"
Write-Host 
# Display Test Mode Status
If (!$Force) {
    Write-Host -NoNewLine -BackgroundColor DarkGreen "### Test Mode ENABLED ###"
	Write-Host " (-Force to Disable)"
}
Else {
    Write-Host -BackgroundColor DarkRed "### Test Mode DISABLED ###"
}

# Pause before begin
Write-Host
Write-Host "Identified"$userList.Count"Users for Updating"
Write-Host
Write-Host
$Ignore = Read-Host “Press ENTER to begin...”
Write-Host

# Prepare Output Tables
$Results = @()
$Successful = @()
$Skipped = @()
$Existing = @()
$Failed = @()

# Commence Updates
ForEach ($userDN in $userList)
{

    # Get AD User Details
    If (!$Credential) {
        $userDetail = Get-ADUser $userDN -Properties Name,Mail,Enabled,ProxyAddresses,Title
    } Else {
        $userDetail = Get-ADUser $userDN -Credential $Cred -Properties Name,Mail,Enabled,ProxyAddresses, Title
    }

    $User = $userDetail.Name

    # Check if Account Disabled
    If (!$userDetail.Enabled) {
        # Account Disabled, add to skipped list
        Write-Host -BackgroundColor DarkYellow ([char]8734) -NoNewline
        Write-Host " $User - Account Disabled, skipping..."
        $Skipped += $User
        Continue
    }
    Write-Debug "  $User - Account enabled, continuing..."

    # Check JobTitle Switch
    If ($JobTitle) {
        Write-Debug "  $User - Job Title check enabled, checking..."
        If ($userDetail.Title -eq $null) {

            # No Job Title Specified, add to skipped list
            Write-Host -BackgroundColor DarkYellow ([char]8734) -NoNewline
            Write-Host " $User - Account missing Job Title, skipping..."
            $Skipped += $User
            Continue
        }
        Write-Debug "  $User - Account Job Title Set, continuing..."

    }
    
    Write-Debug "  $User - Existing Proxy Addresses: $($userDetail.ProxyAddresses)"

    $proxyAddress = "SIP:$($userDetail.mail)"
    Write-Debug "  $User - Proposed Proxy Address is $proxyAddress"

    If ($userDetail.ProxyAddresses -match $proxyAddress) {
        Write-Host -BackgroundColor DarkYellow ([char]8734) -NoNewline
        Write-Host " $User - Account has existing SIP ProxyAddress, skipping..."
        $Existing += $User
        Continue
    }

    # Attempt ProxyAddress Add
    Try {

        # Check Test Mode Status
        If ($Force) {

            # Add new ProxyAddress entry
            Set-ADUser $userDN -add @{ProxyAddresses="$proxyAddress"}
            Write-Host -BackgroundColor DarkGreen ([char]8730) -NoNewline
            Write-Host " $User - Proxy Address added... pending verification..."

            # Update supposedly completed, wait before begin verification
            Start-Sleep 1

            # Re-get AD User Details
            If (!$Credential) {
                $userDetail = Get-ADUser $userDN -Properties Name,Mail,Enabled,ProxyAddresses
            } Else {
                $userDetail = Get-ADUser $userDN -Credential $Cred -Properties Name,Mail,Enabled,ProxyAddresses
            }

            # Verify entry now exists
            If (!$userDetail.ProxyAddresses -match $proxyAddress) {
                Write-Host -BackgroundColor DarkYellow ([char]8734) -NoNewline
                Write-Host " $User - Verify failed, check manually..."
                $Failed += $User
                Continue
            }

            # Verified Change! add to successful list
            Write-Host -BackgroundColor DarkGreen ([char]8730) -NoNewline
            Write-Host " $User - Successful!"
            $Successful += $User

        }
        Else {

            # Simulation Mode, just output to screen
            Write-Host -BackgroundColor DarkYellow ([char]8734) -NoNewline
            Write-Host " $User - Simulate Proxy Address add - $proxyAddress"
            $Successful += "$User*"
            Continue
        }    

    }
    Catch                        
    {

        # Something went wrong, enable debug to show error
        Write-Debug "  $User - $Error[0]"
        Write-Host -BackgroundColor DarkRed ([char]215) -NoNewline
        Write-Host " $User - Error Encountered trying to add Proxy Address, check manually." 
        $Failed += $User
        Continue

    }

}

# Enumerate all results in Table
$FormatEnumerationLimit=-1

# Format Successful Results
$SuccessfulRow = New-Object psobject
$SuccessfulRow | Add-Member -MemberType NoteProperty -Name "Status" -Value "Successful"
$SuccessfulRow | Add-Member -MemberType NoteProperty -Name "Users" -Value $Successful 
$Results += $SuccessfulRow 

# Format Skipped Results
$SkippedRow = New-Object psobject
$SkippedRow | Add-Member -MemberType NoteProperty -Name "Status" -Value "Skipped"
$SkippedRow | Add-Member -MemberType NoteProperty -Name "Users" -Value $Skipped
$Results += $SkippedRow 

# Format Failed Results
$FailedRow = New-Object psobject
$FailedRow | Add-Member -MemberType NoteProperty -Name "Status" -Value "Failed"
$FailedRow | Add-Member -MemberType NoteProperty -Name "Users" -Value $Failed
$Results += $FailedRow

# Format Pre-Existing Results
$ExistingRow = New-Object psobject
$ExistingRow | Add-Member -MemberType NoteProperty -Name "Status" -Value "Pre-Existing"
$ExistingRow | Add-Member -MemberType NoteProperty -Name "Users" -Value $Existing
$Results += $ExistingRow  

If (!$Force) {
    Write-Host
    Write-Host "Note: * denotes simluation User results"
}
# Output Results to Screen
$Results | Format-Table | Out-String | Write-Host



# Stop Transcript
Stop-Transcript | Out-Null

# Pause before Save
$Ignore = Read-Host “Press ENTER to save results to CSV...”
Write-Host

# Output Results to CSV File
$Results | Select-Object Status,@{N='Users';E={$_.Users -Join ','}} | Export-Csv -Path $ResultsFile -NoTypeInformation
Write-Host "Results saved in $ResultsFile"