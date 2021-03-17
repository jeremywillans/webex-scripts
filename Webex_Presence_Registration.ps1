###############################
#
# Webex Presence Integration Registration Script
# Written by Jeremy Willans
# https://github.com/jeremywillans/webex-scripts
# Version: 1.1
#
# USE AT OWN RISK, SCRIPT NOT FULLY TESTED NOR SUPPLIED WITH ANY GURANTEE
#
# Usage - Locates and registers the Webex Outlook DLL, run as Administrator
# Example - powershell -executionpolicy bypass path\to\Webex_Presence_Registration.ps1 <ARGUMENTS>
#
# Parameters:
# -Debug (default false) - Detailed information during process
#
# Change History
# 1.0 20201201 Initial Release
# 1.1 20210317 Formatting Changes
#
###
Param (
    [Switch]$Debug
)
###

# Update Debug Logging
$DebugPreference = "SilentlyContinue"
If ($Debug) {
    $DebugPreference = "Continue"
}

# Confirm Running as Administrator
$IsAdmin = [bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")
If (!$IsAdmin) {
    Write-Host "ERROR - Please run as Administrator, aborting."
    Write-Host 
    Break
}
Write-Debug "Script Running as Administrator, continuing..."

# Define DLL Paths
$Allx64 = "C:\Program Files\Cisco Spark\dependencies\spark-windows-office-integration.dll"
$Allx86 = "C:\Program Files (x86)\Cisco Spark\dependencies\spark-windows-office-integration.dll"
$User = "\AppData\Local\Programs\Cisco Spark\dependencies\spark-windows-office-integration.dll"
$Path = $null

# Test DLL Paths
While (!$Path) {
    # Test X64 Path
    If (Test-Path $Allx64 -PathType Leaf) {
        Write-Debug "MATCH - $($Allx64)"
        $Path = $Allx64
    }
    Write-Debug "NO MATCH - $($Allx64)"

    # Test X86 Path
    If (Test-Path $Allx86 -PathType Leaf) {
        Write-Debug "MATCH - $($Allx86)"
        $Path = $Allx86
    }
    Write-Debug "NO MATCH - $($Allx86)"

    # Test Users Folders
    Get-ChildItem -Directory C:\Users | ForEach-Object {
        $UserPath = "C:\Users\$($_)$($User)"
        If (Test-Path $UserPath -PathType Leaf) {
            Write-Debug "MATCH - $($UserPath)"
            $Path = $UserPath
            Break
        }
        Write-Debug "NO MATCH - $($UserPath)"
    }

    # No Match
    Break
}

If (!$Path) {
    Write-Host "DLL Not Found on System, aborting."
    Write-Host
    Break
}

# Attempt Register DLL
Try {
    # Wrap Path with Quotes
    $Path = """$($Path)"""

    # Output Display if in Debug mode, otherwise use Silent Flag
    If ($Debug) {
        $Process = Start-Process -FilePath 'regsvr32.exe' -Args "$Path" -Wait -NoNewWindow -PassThru
    }
    Else {
        $Process = Start-Process -FilePath 'regsvr32.exe' -Args "/s $Path" -Wait -NoNewWindow -PassThru
    }
    
    # Check Process Exit Code Value
    If ($Process.ExitCode -eq 0) {

        Write-Host "SUCCESS - DLL Registered."
        Write-Host
    }
    Else {

        Write-Host "FAILED to register DLL."
        Write-Host "Run with -Debug for more details"
        Write-Host
    }
}
Catch {
    # Failed
    Write-Host "FAILED to register DLL."
    Write-Host "Run with -Debug for more details"
    Write-Host
    Write-Debug $_.Exception.Message $false
    
}