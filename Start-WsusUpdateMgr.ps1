$VerbosePreference = "Continue"

#Validate user is an Administrator
Write-Verbose "Checking Administrator credentials"
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You are not running this as an Administrator!`nRe-running script and will prompt for administrator credentials."
    Start-Process -Verb "Runas" -File PowerShell.exe -Argument "-STA -noprofile -file $($myinvocation.mycommand.definition)"
	break
}

#Determine if this instance of PowerShell can run WPF 
Write-Verbose "Checking for STA"
If ($Host.Runspace.ApartmentState -ne "STA") {
    Write-Warning "This script must be run in PowerShell started using -STA switch.`nScript will attempt to open PowerShell in STA and run re-run script."
    Start-Process -File PowerShell.exe -Argument "-STA -noprofile -WindowStyle hidden -file $($myinvocation.mycommand.definition)"
    break
}

# Import the Functions
Set-Location $(Split-Path $MyInvocation.MyCommand.Path)
. ".\Functions\Wsus-UpdateManager.ps1"
. ".\UI\Wsus-UpdateManager.UI.ps1"

# Display the GUI
Show-WsusUpdateManagerUi | Out-Null