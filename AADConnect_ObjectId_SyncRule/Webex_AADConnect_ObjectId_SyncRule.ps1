 ################################
#
# AAD Connect Sync Rule Creator for Webex ExternalId (AAD ObjectId)
# Written by Jeremy Willans
# https://github.com/jeremywillans/webex-scripts
# Version: 1.0
#
# USE AT OWN RISK, SCRIPT NOT FULLY TESTED NOR SUPPLIED WITH ANY GURANTEE
#
# Usage - Adds required AAD Connect Rule to sync Azure ObjectID to On-Prem AD for Webex DirSync
# Example - powershell -executionpolicy bypass path\to\Webex_AADConnect_ExternalId_SyncRule.ps1 <ARGUMENTS>
#
# Parameters:
# -Debug (default false) - Detailed information during process
# -Log (default aadsynrule_auditlog.txt) - Appending Transcript Log of process
# -Attribute (default msDS-cloudExtensionAttribute1) - Define AD Attribute to contain AAD ObjectId/ExternalId
# -Precedence (default 148) - Define unique Precedence value for AAD Sync Rule
# -Connector (default Autodiscover) - Define GUID of AD Connector
#
# Change History
# 1.0 20210317 Initial Release
#
###
Param (
    [Switch]$Debug,
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
$ErrorActionPreference = "SilentlyContinue"
Stop-Transcript | Out-Null
$ErrorActionPreference = "Continue"
Start-Transcript -Path $Log -Append

# Clear Screen and Present Script Heading
Clear-Host
Write-Host
Write-Host "AAD Connect Sync Rule for Webex ExternalId"

$adConnector = $null
# Locate Connector for AD
Try {
    If (!$Connector) {
        Write-Debug "Attempting to find AD Connector"
        Get-ADSyncConnector | Where-Object Type -Match "AD" -OutVariable adConnectors | Out-Null
        If ($adConnectors.Count -lt 1) {
            Write-Host "No AD Connector Found!"
            Exit 1
        }
        If ($adConnectors.Count -ne 1) {
            Write-Host "ERROR: Multiple AD Connectors identified, please define with script execution ( -Connector <Identifier>)"
            $adConnectors | Select-Object Name, Identifier | Format-Table | Out-String | Write-Host
            Exit 1
        }

        # Set Connector
        $adConnector = $adConnectors[0]
        Write-Debug "Connector Found!"

    }
    Else {

        Write-Debug "Using provided AD Connector Identifier"
        # Define Connector, this will fail to Catch if invalid
        $adConnector = Get-ADSyncConnector -Identifier $Connector
        Write-Debug "Connector Found!"
    }
}
Catch {
    Write-Host "ERROR: Unable to locate connector!"
    Write-Host $_.Exception.Message
    Write-Debug $_.Exception
    Exit 1
}

# Check Precedence
Write-Debug "Checking Precedence"
Try {
    $PrecedenceCheck = Get-ADSyncRule | Where-Object Precedence -EQ $Precedence
    If ($PrecedenceCheck.Count -ne 0) {
        Write-Host "ERROR: Precedence value of $Precedence is not unique, please define during script execution with -Precedence <Number>"
        Exit 1
    }
    Write-Debug "Precedence Unique!"
}
Catch {
    Write-Host "Error Checking Precedence"
    Write-Host $_.Exception.Message
    Write-Debug $_.Exception
    Exit 1

}

# Confirm Running as Administrator
$IsAdmin = [bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")
If (!$IsAdmin) {
    Write-Host 
    Write-Host "ERROR - Please run as Administrator, aborting."
    Write-Host 
    Exit 1
}

# Output options
Write-Host
Write-Host "Log File: $Log"
Write-Host "AD Connector:"$adConnector.Name
Write-Host "Destination Attribute: $Attribute"
Write-Host "Precedence: "$Precedence
Write-Host

# Pause before begin
Write-Host
Read-Host "Press ENTER to begin..."
Write-Host

Try {
    # Create AD Sync Rule
    Write-Debug "Create AD Sync Rule"
    New-ADSyncRule  `
        -Name 'Out to AD - User AAD ObjectId' `
        -Description 'Maps AAD ObjectId to AD Attribute for Webex' `
        -Direction 'Outbound' `
        -Precedence $Precedence `
        -PrecedenceAfter '00000000-0000-0000-0000-000000000000' `
        -PrecedenceBefore '00000000-0000-0000-0000-000000000000' `
        -SourceObjectType 'person' `
        -TargetObjectType 'user' `
        -Connector $adConnector.Identifier.Guid `
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
        -ArgumentList 'sourceAnchor', '', 'ISNOTNULL' `
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
    If ($Debug) {
        # Show Full Rule in Debug
        Get-ADSyncRule  `
        -Identifier $NewRule.Identifier
    } Else {
        Get-ADSyncRule  `
            -Identifier $NewRule.Identifier | Select-Object Identifier, Name, Description
    }
}
Catch {
    Write-Host "Error Encountered"
    Write-Host $_.Exception.Message
    Write-Debug $_.Exception
    Exit 1
}

# Stop Transcript
Stop-Transcript | Out-Null