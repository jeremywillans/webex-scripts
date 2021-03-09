###############################
#
# Webex Teams Create Email Address Script
# Written by Jeremy Willans, Cisco Systems Australia.
# Date: 17-11-2020
#
# USE AT OWN RISK, SCRIPT NOT FULLY TESTED NOR SUPPLIED WITH ANY GURANTEE
#
# Usage - Include in startup GPO to remove requirement for users to enter email Address when WxT loads.
# Example - powershell -executionpolicy bypass path\to\Webex_Create_Email_Alias.ps1
#
###

# Define Static Variables 
$registryPath = "HKCU:\Software\Cisco Spark Native\Params"
$name = "email"

Try {

    # Attempt Find Current User Email
    $searcher = [adsisearcher]"(samaccountname=$env:USERNAME)"
    $email = $searcher.FindOne().Properties.mail
} Catch {
    
    # Unable to reach Domain
    Write-Host "Domain not reachable, aborting."
    Break
}

# Check if Email is defined
If ($email -eq $null) {
    Write-Host "No Email Address found, aborting."
    Break
}

Try {

    # Check if Reg Path Exists
    If(!(Test-Path $registryPath))
    {
        # Create Reg Path
        New-Item -Path $registryPath -Force | Out-Null
    }

    # Create/Update Email Reg Key
    New-ItemProperty -Path $registryPath -Name $name -Value $email -PropertyType String -Force | Out-Null
    Write-Host "Done."

} Catch {

    # Something went wrong.
    Write-Host "Failed to add/update Email, check manually."
}