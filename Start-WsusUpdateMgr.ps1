$VerbosePreference = "Continue"

#Validate user is an Administrator
Write-Verbose "Checking Administrator credentials"
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You are not running this as an Administrator!`nRe-running script and will prompt for administrator credentials."
    Start-Process -Verb "Runas" -File PowerShell.exe -Argument "-STA -noprofile -file $($myinvocation.mycommand.definition)"
	    Break
}

# Import the Functions
. ".\Functions\Wsus-UpdateManager.ps1"
. ".\UI\Wsus-UpdateManager.UI.ps1"

# Display the GUI
Show-WsusUpdateManagerUi | Out-Null