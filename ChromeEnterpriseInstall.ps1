<#	
	.NOTES
#===========================================================================
	Created by:   	Ron Knight
	Date:           01.08.2020
	Organization: 	##############
	Version:        1.0
	Filename:    Install.ps1
#===========================================================================
	.DESCRIPTION
		This installs Google Chrome v. 79.0.3945.117
	    Sourcefile can be obtained from:
        https://cloud.google.com/chrome-enterprise/browser/download/
	Revision Notes:
#>

#===========================================================================
#Set Variables for script
$AppName = "Chrome79.0.3945.117" #Name of application as it will appear in the log
$File = "GoogleChromeStandaloneEnterprise64.msi"
$LogPath = "C:\temp\BFSlogs" #location of log file
$LogDate = get-date -format "MM-dd-yyyy"
$OS = Get-WmiObject Win32_OperatingSystem
$OSCap = $OS.Caption
$Arch = $OS.OSArchitecture
$OSBuild = $OS.BuildNumber
$ComputerSystem = Get-WmiObject Win32_ComputerSystem
if ($ComputerSystem.Manufacturer -like 'Lenovo') { $Model = (Get-WmiObject Win32_ComputerSystemProduct).Version }
            else { $Model = $ComputerSystem.Model }
$LastLogon = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\').LastLoggedOnUser


#===========================================================================
#Begin Logging Script
if (!(test-path C:\temp\BFSlogs))
{
    Write-Output "Log folder for BFS does not exist - creating it."
	New-Item -ItemType Directory -Path C:\temp\BFSlogs -Verbose
	Start-Transcript -Path "$LogPath\$AppName.$LogDate.log" -Force -Append
}

else { 
	Start-Transcript -Path "$LogPath\$AppName.$LogDate.log" -Force -Append
}

#===========================================================================
Write-Output "### Script Start ###"
Write-Output "Start time: $(Get-Date)"
Write-Output "Username: $(([Environment]::UserDomainName + "\" + [Environment]::UserName))"
Write-Output "Last Logged On User: $lastlogon"
Write-Output "Computer Name: $(& hostname)"
Write-Output "Operating System: $OSCap $Arch"
Write-Output "Build = $OSBuild"
Write-Output "Computer Model: $Model"
#===========================================================================
Write-Output "Beginning Silent Install of $AppName MSI"

Start-Process -filepath msiexec -ArgumentList "/i $File /q" -Wait -PassThru
Write-Output "$AppName install complete."

#===========================================================================

Write-Output "Modifying Registry to disable AutoUpdate"

New-Item -Path "HKLM:\SOFTWARE\Policies\Google" -Name \Update –Force

Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Google\Update" -Name “AutoUpdateCheckPeriodMinutes” -Value 0
Write-Output "AutoUpdateCheckPeriodMinutes value set to 0"

Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Google\Update" -Name “DisableAutoUpdateChecksCheckboxValue” -Value 0
Write-Output "DisableAutoUpdateChecksCheckboxValue value set to 0"

Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Google\Update" -Name “UpdateDefault” -Value 0
Write-Output "UpdateDefault value set to 0"

Write-Output "All registry entries created."
#===========================================================================

Write-Output "$AppName script installation Complete."
# ===========================================================================
Stop-Transcript
