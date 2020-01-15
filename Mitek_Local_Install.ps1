<#	
	.NOTES
#===========================================================================
	Created by:   	Wesley Sinnen
    Modified by:    Ron Knight
    Date:           11.08.2019
	Organization: 	Builders First Source
	Filename:    MiTek_Local_Install.ps1
#===========================================================================
	.DESCRIPTION
		This installs Mitek 8.3.2 with an option to upgrade existing versions.
#>


#Start-Transcript -Path "C:\temp\BFSlogs\Mitek_Upgrade8.3.2_LocalInstall.log" -Append -Force
#Write-Output "Start"
$MitekSoftwareRepository  = "\\dalsccmpri01\PKGSource\Mitek\Mitek-v8.3.2 GR"
$MitekRemoteServerPath = "\Mitek$\v8.3.2\Mitek-v8.3.2 GR"
$MitekTempFolder ="C:\Temp\MiTek-Install"
$LogDate = get-date -format "MM-dd-yy-HH"

function Get-MII-DB{
    param (
        [string]$MitekPath,
        [string]$MitekVersion
    ) 
      [hashtable]$return = @{}
     $MitekINI = Get-Content -Path $MitekPath\Programs\Mitek.ini
     $MitekINI | ForEach-Object {
         If($_ -Like '*Current=*'){
             $CurrentConfig = $_.trim().trim("Current=") 
         }
     }

     $lstStatus.Items.Add("Current Config: $CurrentConfig")

     $MitekINI | ForEach-Object {
         If($_ -Match  "\[$CurrentConfig\]"){
             $FoundConfig = 1
             $lstStatus.Items.Add("Found Current Config in Mitek.INI $_")
         }  
         If($FoundConfig -Match 1 -and $_ -Like "*DBDefault=*"){
             $FoundConfig = 2
             $DBConfig = $_.trim("DBDefault=")
             $lstStatus.Items.Add($DBConfig.Trim())
             $DBSourceTemp, $DBCatalogTemp, $DBUserIDTemp, $DBPasswordTemp, $DBTimeoutTemp = $DBConfig.split(";")
             $DBSource = $DBSourceTemp.trim().trim("Source=")
             $DBCatalog = $DBCatalogTemp.trim().trim("Initial Catalog=")
         }
     }
     $return.DBSource = $DBSource
     $return.DBCatalog = $DBCatalog
     return $return
}

function msgbox {
    param (
        [string]$Message,
        [string]$Title = 'Message box title',   
        [string]$buttons = 'OKCancel'
    )
    # This function displays a message box by calling the .Net Windows.Forms (MessageBox class)
 
    # Load the assembly
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
 
    # Define the button types
    switch ($buttons) {
       'ok' {$btn = [System.Windows.Forms.MessageBoxButtons]::OK; break}
       'okcancel' {$btn = [System.Windows.Forms.MessageBoxButtons]::OKCancel; break}
       'AbortRetryIgnore' {$btn = [System.Windows.Forms.MessageBoxButtons]::AbortRetryIgnore; break}
       'YesNoCancel' {$btn = [System.Windows.Forms.MessageBoxButtons]::YesNoCancel; break}
       'YesNo' {$btn = [System.Windows.Forms.MessageBoxButtons]::yesno; break}
       'RetryCancel'{$btn = [System.Windows.Forms.MessageBoxButtons]::RetryCancel; break}
       default {$btn = [System.Windows.Forms.MessageBoxButtons]::RetryCancel; break}
    }
 
    # Display the message box
    $Return=[System.Windows.Forms.MessageBox]::Show($Message,$Title,$btn)
    $Return
}

function ExitScript {
    #Stop Transcript
    Stop-Transcript
    $Form.Close()
    }

function GetCurrentLocation {
     [hashtable]$return = @{}
     $lstStatus.Items.Add("Trying to determin what location you are at.")
     $SiteInfo = Get-Content -Path SiteInfo.ini
     $IPAddress = (Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter 'IPEnabled = True').IPAddress
     foreach ($element in $IPAddress) {
        #$lstStatus.items.Add("IP Search -------------")
        #$lstStatus.items.Add($element)
        $SiteInfo | ForEach-Object {
             $SiteLocationTemp, $IPSubNetTemp, $SiteServerTemp = $_.Split(",")
             
             #$lstStatus.items.Add($IPSubNetTemp)
             
             if ($element -like "$IPSubNetTemp*") {
                  
                  #$lstStatus.items.Add($IPSubNetTemp)
                  $return.SiteLocation = $SiteLocationTemp
                  $return.IPSubNet = $IPSubNetTemp
                  $return.SiteServer = $SiteServerTemp
                  return $return
                  break
             }
        }
     }
     if ([string]::IsNullOrEmpty($return.SiteLocation)) {
         $return.SiteLocation = "Remote"
         $return.IPSubNet = "10.10.20"
         $return.SiteServer = "Dallas"
         return $return
        }

}

function DownloadMitekInstall {
     param (
        [string]$SiteLocation,
        [string]$SiteServer
     ) 

     if ($SiteLocation -like "Remote") 
          {
          $lstStatus.Items.Add("You seem to be remote to a truss plant.  Copying install from Dallas." ) 
          }
     else {
        $MitekSoftwareRepository = "\\$SiteServer$MitekRemoteServerPath"
        }
     $lstStatus.items.Add("Downloading Mitek install from $MitekSoftwareRepository ")
     $lstStatus.Items.Add("Destination Folder:  $MitekTempFolder")
     $lstStatus.Items.Add("Downloading Mitek install.  Please wait. This part can take a while." ) 

     Robocopy $MitekSoftwareRepository $MitekTempFolder /MIR
     }


function CheckForMitekFolder {
    param (
        [string]$MitekPath = "C:\Mitek"
    )
    if (test-path "$MitekPath\Programs\ou.exe") {
        return $True
    {
    Else
    }
        return $False
    }
}

Function FreshInstall {
    Start-Transcript -Path "C:\temp\BFSlogs\Mitek 8.3.2_$LogDate.log" -Append -Force
    #Set-Location Programs\InitialSetup

    $lstStatus.Items.Add("Starting Fresh Install:") 
    $lstStatus.Items.Add("$File") 
    $lstStatus.Items.Add("$ARG") 
    Start-Process "$MitekTempFolder\Programs\InitialSetup\Setup.exe" -ArgumentList '/z "MBA=true;Path=C:\MiTek;Folder=MiTekSuite;SQLName=(local)\MII;DBName=MII;"' -wait
}

if (!(test-path C:\temp\BFSlogs)){
	New-Item -ItemType Directory -Path C:\temp\BFSlogs -Verbose
}

$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
#Write-Output $scriptPath
Set-Location $scriptPath

$RegistryPath = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
Set-Location "$RegistryPath"
$allkeys = Get-ChildItem

$mitekinstallpath = New-Object System.Collections.ArrayList
$mitekinstallversion = New-Object System.Collections.ArrayList
[boolean]$foundInstall = $false

Foreach ($item in $allkeys)
{
	$regkey = Get-ItemProperty -Path ($item -replace ("hkey_local_machine", "HKLM:")) -Name "displayname", "installlocation", "displayversion" -ErrorAction SilentlyContinue -Verbose
	if ($regkey.displayname -like "MiTek SAPPHIRE Structure" + ("*"))	{
        $mitekinstallpath.Add($regkey.installlocation ) > $null
        $mitekinstallversion.Add($regkey.DisplayVersion ) > $null
        #Write-Output $regkey.installlocation
        #Write-Output $regkey.DisplayVersion
        $foundInstall = $true
	} 	
}

Set-Location $scriptPath
#Write-Output $scriptPath

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '600,600'
$Form.text                       = "Mitek Install / Upgrade Menu"
$Form.TopMost                    = $false

$lstMitekInstallPath                        = New-Object system.Windows.Forms.ListBox
$lstMitekInstallPath.text                   = "listBox"
$lstMitekInstallPath.width                  = 100
$lstMitekInstallPath.height                 = 271
$lstMitekInstallPath.location               = New-Object System.Drawing.Point(15,20)
$lstMitekInstallPath.Items.AddRange($mitekinstallpath)

$lstMitekInstallVersion                        = New-Object system.Windows.Forms.ListBox
$lstMitekInstallVersion.text                   = "listBox"
$lstMitekInstallVersion.width                  = 130
$lstMitekInstallVersion.height                 = 271
$lstMitekInstallVersion.location               = New-Object System.Drawing.Point(115,20)
$lstMitekInstallVersion.Items.AddRange($mitekinstallversion)

$lstStatus                        = New-Object system.Windows.Forms.ListBox
$lstStatus.text                   = "listBox"
$lstStatus.width                  = 570
$lstStatus.height                 = 271
$lstStatus.location               = New-Object System.Drawing.Point(15,300)

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "New Install"
$Button1.width                   = 130
$Button1.height                  = 30
$Button1.location                = New-Object System.Drawing.Point(250,15)
$Button1.Font                    = 'Microsoft Sans Serif,10'

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.text                    = "Upgrade Selected"
$Button2.width                   = 130
$Button2.height                  = 30
$Button2.location                = New-Object System.Drawing.Point(250,50)
$Button2.Font                    = 'Microsoft Sans Serif,10'

$Form.controls.AddRange(@($lstMitekInstallPath,$lstMitekInstallVersion,$lstStatus,$Button1,$Button2))

$lstMitekInstallPath.Add_SelectedIndexChanged(
       {
       $lstMitekInstallVersion.SelectedIndex = $lstMitekInstallPath.SelectedIndex
	   }
	)

$lstMitekInstallVersion.Add_SelectedIndexChanged(
       {
        $lstMitekInstallPath.SelectedIndex = $lstMitekInstallVersion.SelectedIndex
		}
	)

$Button1.Add_Click(  
        {
        $lstStatus.Items.Add("You have choosen to install a new instance of Mitek")
        #$lstStatus.Items.Add("Checking for an existing Mitek folder")
            If (CheckForMitekFolder) {
               MsgBox -Message "I detected there is a Mitek folder that is not empty"
               ExitScript  
           } Else {
               #$lstStatus.Items.Add("No Mitek folders detected")
               $CurrentLocation = GetCurrentLocation

               $lstStatus.Items.Add($CurrentLocation.SiteLocation)
               $lstStatus.Items.Add($CurrentLocation.IPSubNet)
               $lstStatus.Items.Add($CurrentLocation.SiteServer)
            
               DownloadMitekInstall -SiteLocation $CurrentLocation.SiteLocation -SiteServer $CurrentLocation.SiteServer
               FreshInstall
           }
           [System.Environment]::Exit(0)
        }
    )	

$Button2.Add_Click(  
        {
        $Path2Upgrade = $lstMitekInstallPath.SelectedItem
        $Version2Upgrade = $lstMitekInstallVersion.SelectedItem
        
        if ($Version2Upgrade -like "7*")
            {
            $lstStatus.Items.Add("Mitek 7.6.4 can’t be directly upgraded.  Please upgrade it manually")
            msgbox -Title "Warning" -buttons 'Ok' -Message "Mitek 7.6.4 can’t be directly upgraded.  Please upgrade it manually"
            }
        elseif ($Version2Upgrade -like "8*")
            {
            $File = "$MitekTempFolder\SetupCMD"
            $ARG = $lstMitekInstallPath.SelectedItem
            
            $CurrentLocation = GetCurrentLocation

            $lstStatus.Items.Add($CurrentLocation.SiteLocation)
            $lstStatus.Items.Add($CurrentLocation.IPSubNet)
            $lstStatus.Items.Add($CurrentLocation.SiteServer)
            $SiteServer = $CurrentLocation.SiteServer

            $DBSettings = Get-MII-DB -MitekPath $Path2Upgrade -MitekVersion $Version2Upgrade
            
            if ($DBSettings.DBSource -like "*$SiteServer*") {
                 msgbox -Title "Warning" -buttons 'Ok' -Message "It looks like you are configured to a shared database.  Please set you database local and restart the upgrade."
                 [System.Environment]::Exit(0)
                 }
            

            DownloadMitekInstall -SiteLocation $CurrentLocation.SiteLocation -SiteServer $CurrentLocation.SiteServer
            $lstStatus.Items.Add("Starting upgrade on Mitek $Version2Upgrade") 
            Start-Process -FilePath $File -ArgumentList $ARG -Wait
            [System.Environment]::Exit(0)
            }
        else 
            {
            $lstStatus.Items.Add("The versoin selected is not reconized")
            ExitScript
            }
              
		}
    )	

If ($foundInstall) {
     #[void]$Form.ShowDialog()
     $Form.ShowDialog() | Out-Null
     }
ElseIf (CheckForMitekFolder) {
    MsgBox -Title "Warning" -buttons 'Ok' -Message "There alreade seems to be a Mitek folder with programs.  Clean install will not run untill this has been resovled"
    }
Else 
    {
    $Button2.Visible                 = $False
    $lstMitekInstallPath.Visible     = $False
    $lstMitekInstallVersion.Visible  = $False

    $lstStatus.location               = New-Object System.Drawing.Point(15,50)
    $lstStatus.height                 = 531


     #[void]$Form.ShowDialog()
     $Form.ShowDialog() | Out-Null
    } 
