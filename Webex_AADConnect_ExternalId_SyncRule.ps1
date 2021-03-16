################################
#
# AAD Connect Sync Rule Creator for Webex ExternalId
# Written by Jeremy Willans, Cisco Systems Australia.
# Date: 17-03-2021
#
# USE AT OWN RISK, SCRIPT NOT FULLY TESTED NOR SUPPLIED WITH ANY GURANTEE
#
# Parameters:
#
# -Debug (default false) - Detailed information during process
# -Log (default aadsynrule_auditlog.txt) - Appending Transcript Log of process
# -Attribute (default msDS-cloudExtensionAttribute1) - Define AD Attribute to contain AAD ObjectId/ExternalId
# -Precedence (default 148) - Define unique Precedence value for AAD Sync Rule
# -Connector (default Autodiscover) - Define GUID of AD Connector
#
# Usage - Adds required AAD Connect Rule to sync Azure ObjectID to On-Prem AD for Webex DirSync
# Example - powershell -executionpolicy bypass path\to\Webex_AADConnect_ExternalId_SyncRule.ps1 <ARGUMENTS>
#
###
Param (
	[Switch]$Debug = $true,
    [String]$Log = "$PSScriptRoot\aadsynrule_auditlog.txt",
    [String]$Attribute = "msDS-cloudExtensionAttribute1",
    [Int]$Precedence = 148,
    [String]$Connector
 )

# Update Debug Logging
$DebugPreference = "SilentlyContinue"
If ($Debug) {
    $DebugPreference = "Continue"
}

# Setup Transcript
$ErrorActionPreference="SilentlyContinue"
Stop-Transcript | Out-Null
$ErrorActionPreference = "Continue"
Start-Transcript -path $Log -append

# Clear Screen and Present Script Heading
cls
Write-Host
Write-Host "AAD Connect Sync Rule for Webex ExternalId"

# Locate Connector for AD
Try
{
    If ($Connector -eq "") {
        Write-Debug "Attempting to find AD Connector"
        $Connectors = Get-ADSyncConnector | Where Type -Match "AD" #-OutVariable Connectors | Out-Null
        If ($Connectors.Count -lt 1) {
            Write-Host "No AD Connector Found!"
            Break
        }
        If ($Connectors.Count -ne 1) {
            Write-Host "ERROR: Multiple AD Connectors identified, please define with script execution ( -Connector <Identifier>)"
            $Connectors | Select-Object Name, Identifier | Format-Table | Out-String | Write-Host
            Break
        }

        # Set Connector
        $Connector = $Connectors[0]
        Write-Debug "Found!"

    } Else {

        $Connector = Get-ADSyncConnector -Identifer $Connector
        If ($PrecedenceCheck.Count -ne 0) {
            Write-Host "ERROR: Invalid Connector Identifier Specified."
            Break
        }
    }
}
Catch
{
    Write-Host "ERROR: Unable to locate connector!"
    Write-Host $Error[0]
    Write-Debug $Error
    Break
}

# Check Precedence
Write-Debug "Checking Precedence"
Try
{
    $PrecedenceCheck = Get-ADSyncRule | Where Precedence -eq $Precedence
    If ($PrecedenceCheck.Count -ne 0) {
        Write-Host "ERROR: Precednce is not unique, please define during script execution with -Precedence <Number>"
        Break
    }
    Write-Debug "Precedence Unique!"
}
Catch
{
    Write-Host "Error Checking Precedence"
    Write-Host $Error[0]
    Write-Debug $Error
    Break
}


# Confirm Running as Administrator
$IsAdmin = [bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")
If (!$IsAdmin) {
    Write-Host 
    Write-Host "ERROR - Please run as Administrator, aborting."
    Write-Host 
    Break
}

# Output options
Write-Host
Write-Host "Log File: $Log"
Write-Host "AD Connector:"$Connector.Name
Write-Host "Destination Attribute: $Attribute"
Write-Host "Precedence: "$Precedence
Write-Host

# Pause before begin
Write-Host
$Ignore = Read-Host “Press ENTER to begin...”
Write-Host

Try
{

    # Create AD Sync Rule
    Write-Debug "Create AD Sync Rule"
    New-ADSyncRule  `
    -Name 'Out to AD - User ExternalId for Webex' `
    -Description '' `
    -Direction 'Outbound' `
    -Precedence $Precedence `
    -PrecedenceAfter '00000000-0000-0000-0000-000000000000' `
    -PrecedenceBefore '00000000-0000-0000-0000-000000000000' `
    -SourceObjectType 'person' `
    -TargetObjectType 'user' `
    -Connector $Connector.Identifier.Guid `
    -LinkType 'Join' `
    -SoftDeleteExpiryInterval 0 `
    -ImmutableTag '' `
    -OutVariable syncRule | Out-Null

    # Create Attribute Mapping for ExternalId
    Write-Debug "Create Attribute Mapping"
    Add-ADSyncAttributeFlowMapping  `
    -SynchronizationRule $syncRule[0] `
    -Destination $Attribute `
    -FlowType 'Expression' `
    -ValueMergeType 'Update' `
    -Expression 'Replace([cloudAnchor],"User_","")' `
    -OutVariable syncRule | Out-Null

    # Create Scope Condition
    Write-Debug "Create Scope Condition"
    New-Object  `
    -TypeName 'Microsoft.IdentityManagement.PowerShell.ObjectModel.ScopeCondition' `
    -ArgumentList 'sourceAnchor','','ISNOTNULL' `
    -OutVariable condition0 | Out-Null

    # Add Scope Condition to Rule
    Write-Debug "Scope Condition to Rule"
    Add-ADSyncScopeConditionGroup  `
    -SynchronizationRule $syncRule[0] `
    -ScopeConditions @($condition0[0]) `
    -OutVariable syncRule | Out-Null 

    Write-Debug "Add Rule"
    # Add ADSync Rule
    Add-ADSyncRule  `
    -SynchronizationRule $syncRule[0] -OutVariable NewRule | Out-Null

    Write-Host "Rule Created!"

    # Display Result
    Get-ADSyncRule  `
    -Identifier $NewRule.Identifier

}
Catch
{
    Write-Host "Error Encountered"
    Write-Host $Error[0]
    Write-Debug $Error
}