###############################
#
# Webex App Create Email Address Script
# Written by Jeremy Willans
# https://github.com/jeremywillans/webex-scripts
# Version: 1.1
#
# USE AT OWN RISK, SCRIPT NOT FULLY TESTED NOR SUPPLIED WITH ANY GURANTEE
#
# Usage - Include in startup GPO to remove requirement for users to enter email Address when WxApp loads.
# Example - powershell -executionpolicy bypass path\to\Webex_Create_Email_Alias.ps1
#
# Change History
# 1.0 20201117 Initial Release
# 1.1 20210317 Formatting Changes
#
###

# Define Static Variables 
$registryPath = "HKCU:\Software\Cisco Spark Native\Params"
$name = "email"

Try {

    # Attempt Find Current User Email
    $searcher = [adsisearcher]"(samaccountname=$env:USERNAME)"
    $email = $searcher.FindOne().Properties.mail
}
Catch {
    
    # Unable to reach Domain
    Write-Host "Domain not reachable, aborting."
    Break
}

# Check if Email is defined
If (!$email) {
    Write-Host "No Email Address found, aborting."
    Break
}

Try {

    # Check if Reg Path Exists
    If (!(Test-Path $registryPath)) {
        # Create Reg Path
        New-Item -Path $registryPath -Force | Out-Null
    }

    # Create/Update Email Reg Key
    New-ItemProperty -Path $registryPath -Name $name -Value $email -PropertyType String -Force | Out-Null
    Write-Host "Done."

}
Catch {

    # Something went wrong.
    Write-Host "Failed to add/update Email, check manually."
}