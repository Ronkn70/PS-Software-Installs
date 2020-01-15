<#	
	.NOTES
#===========================================================================
	Created by:   	Ron Knight
	Date:          01.02.2020
	Organization: 	Builders First Source
	Version:        1.0.0
	Filename:    Mitek_Analytics.ps1
#===========================================================================
	.DESCRIPTION
		This script pulls the Mitek GUID 
        and compiles it into a CSV with the Computername and GUID
	
	Revision Notes:
#>

#===========================================================================
#Set Variables for script
$AppName = "Mitek_GUID" #Name of application as it will appear in the log
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
$CSVPath = "C:\Working\Mitek Analytics"
#===========================================================================
#Begin Logging Script
if (!(test-path C:\temp\BFSlogs))
{
    Write-Output "Log folder for BFS does not exist - creating it."
	New-Item -ItemType Directory -Path C:\temp\BFSlogs -Verbose
    }
else { }
Start-Transcript -Path "$LogPath\$AppName.$LogDate.log" -Force -Append

#===========================================================================
Write-Output "### Script Start ###"
Write-Output "Start time: $(Get-Date)"
Write-Output "Username: $(([Environment]::UserDomainName + "\" + [Environment]::UserName))"
Write-Output "Last Logged On User: $lastlogon"
Write-Output "Computer Name: $(& hostname)"
Write-Output "Operating System: $OSCap"
Write-Output "Architecture: $Arch"
Write-Output "Build = $OSBuild"
Write-Output "Computer Model: $Model"
#===========================================================================
$users = get-childitem c:\users
foreach ($user in $users) {
    if (Test-Path C:\Users\$user\AppData\Local\MiTek\Telemetry\Analytics.cfg -PathType Leaf) { 
        $json = Get-Content 'C:\Users\$user\AppData\Local\MiTek\Telemetry\Analytics.cfg' | Out-String | ConvertFrom-Json
        Write-Output $json.UserId    
        $Properties =@{
            UserName = "$User"
            GUID = "$json.UserId"
            }
        $o = New-Object psobject -Property $Properties; $O
        $o | Export-Csv "$CSVPath" -app
        }
        Else {
        Write-Output "No Mitek telemetry data for $user"
        }
}
# ===========================================================================
Stop-Transcript