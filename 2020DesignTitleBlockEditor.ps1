<#	
	.NOTES
	===========================================================================
	Created by:   	Ron Knight
	Organization: 	#############
	Filename:     	install.ps1
    ===========================================================================
	.DESCRIPTION
		THis installs 2020 Title Block Editor 3.3.14.2
#>

#Create program folders in program files (x86)

$destinationFolder = "C:\Program Files (x86)\2020 Title Block Editor\3.3.14.2"
if (!(Test-Path -path $destinationFolder)) {New-Item $destinationFolder -Type Directory}
Copy-Item ".\*" -Destination $destinationFolder -Recurse -Force

#Create desktop shortcut in Allusers desktop

$desktop = "C:\Users\Public\Desktop"
Copy-Item ".\*.lnk" -Destination $desktop -Verbose
#exit 0