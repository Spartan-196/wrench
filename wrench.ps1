
###############################################################################
# Program Name: Wrench
# Author: Matt Tuchfarber
# Contributors: James Atkinson, Kevin Cook
# Code Refactored by: Kevin Cook 
# Current maintainer: James Atkinson
# Date Created: 2014-09-12
# Purpose: To enable quicker searching of users and computers in active directory
#			and to allow easy manipulation of those objects. A toolkit built
#			to ease help desk support and work more efficiently 
###############################################################################
#Load Vairables from Enviroment File
. "$PSScriptRoot\wrench_env.ps1"

### First two functions force script to be ran as admin ###
function IsAdministrator{
    $Identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $Principal = New-Object System.Security.Principal.WindowsPrincipal($Identity)
    $Principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

function IsUacEnabled{ (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System).EnableLua -ne 0 }

if (!(IsAdministrator)){
    if (IsUacEnabled){	

        [string[]]$argList = @('-NoProfile','-WindowStyle','-File', $MyInvocation.MyCommand.Path)
        $argList += $MyInvocation.BoundParameters.GetEnumerator() | ForEach-Object {"-$($_.Key)", "$($_.Value)"}
        $argList += $MyInvocation.UnboundArguments
        Start-Process PowerShell.exe -Verb Runas -WorkingDirectory $pwd -ArgumentList $argList 
        return
    }else{
        throw "You must be administrator to run this script"
    }
}
############################################################

###Load Windows Form Assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")   
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")   
[System.Windows.Forms.Application]::EnableVisualStyles();

$msg = [System.Windows.Forms.MessageBox]	# Standard message box form

### Optimal Prereq checks
$correctPSVersion = $false
$rsatInstalled = $false
$sccmConfigPSD1 = test-path "$SCCMSiteDataFile"
#check for PSd1 file that allows powershell to connect to SCCM Site Server for cmdlets


#Check Powershell version > 3
if ($psversiontable.psversion.Major -lt 3){
	$msg::Show("This program requires Powershell version 3 or higher")
}else{
	$correctPSVersion = $true
}

if ($correctPSVersion -eq $True){
#So long as Powesehell version is good, look for AD module of RSAT
	if(get-module -list activedirectory){
	$rsatInstalled=$true
	}else{$msg::Show("This program requires Microsoft RSAT to be installed")
#Exit if RSAT is not installed as Wrench cannot be used without it.
	}
}

if ($correctPSVersion -eq $true -AND $rsatInstalled -eq $true){
	Import-Module ActiveDirectory
	
	If ($sccmConfigPSD1 -eq $True) {
	Import-Module $SCCMSiteDataFile #import sccm data file
	$CMDrive = (get-psdrive -PSProvider CMSite).Name
	$CMPSSuppressFastNotUsedCheck = $true #Suppress warning messages about -fast being supported in some instances
	} 
	
	#Used to created GUI items
	function createItem($Type, $LocationX, $LocationY, $SizeX, $SizeY, $Text, $ParentForm){
		$template = New-Object System.Windows.Forms.$Type
		$template.Location = New-Object System.Drawing.Size($LocationX,$LocationY)
		$template.Size = New-Object System.Drawing.Size($SizeX,$SizeY)
		$template.Text = $Text
		$ParentForm.Controls.Add($template)
		$template
	}
	
	#Used to create GUI forms
	function createForm($Text, $SizeX, $SizeY, $StartPos, $BorderStyle, $MinBox, $MaxBox, $ControlBox){
		$template = New-Object System.Windows.Forms.Form
		$template.Text = $Text
		$template.Size = New-Object System.Drawing.Size($SizeX,$SizeY)
		$template.StartPosition = $StartPos
		$template.MinimizeBox = $MinBox
		$template.MaximizeBox = $MaxBox
		$template.ControlBox = $ControlBox
		$template.FormBorderStyle = $BorderStyle
		$template
	}
	
#Used to display a created form
function showForm($Form){
	$Form.Add_Shown({$Form.Activate()})	
	[void] $Form.ShowDialog()
	$Form.BringToFront()	
}

####### MAIN FORM GUI #######
$mainForm = createForm "Wrench" 300 635 "CenterScreen" "Fixed3D" $true $false $true
	$mainForm.Icon = New-Object system.drawing.icon($IconLocation)
	$mainForm.KeyPreview = $True
	$mainForm.Add_KeyDown({ mainKeyboard })
#Name Info
$UserNameLabel									= New-Object system.Windows.Forms.Label
$UserNameLabel.text								= "Name: "
$UserNameLabel.width 							= 55
$UserNameLabel.height							= 20
$UserNameLabel.location							= New-Object System.Drawing.Point(11,28)
$UserNameLabel.Font                         	= 'Microsoft Sans Serif,8.25'

$UserNameTextbox                        	  	= New-Object system.Windows.Forms.TextBox
$UserNameTextbox.width                        	= 130
$UserNameTextbox.height                       	= 20
$UserNameTextbox.location                     	= New-Object System.Drawing.Point(70,25)
$UserNameTextbox.Font                         	= 'Microsoft Sans Serif,8.25'

$UserNameSearchButton                           	= New-Object system.Windows.Forms.Button
$UserNameSearchButton.text                      = "Search"
$UserNameSearchButton.width                     = 60
$UserNameSearchButton.height                    = 20
$UserNameSearchButton.location                  = New-Object System.Drawing.Point(210,25)
$UserNameSearchButton.Font                      = 'Microsoft Sans Serif,8.25'
$UserNameSearchButton.TabStop = $False
$UserNameSearchButton.Add_Click({	searchByName })
# $UserNameSearchButton = createItem "Button" 210 25 60 20 "Search" $mainForm
	

$UserIDLabel                            		= New-Object system.Windows.Forms.Label
$UserIDLabel.text                      			= "User ID: "
$UserIDLabel.width                      		= 60
$UserIDLabel.height                     		= 20
$UserIDLabel.location                   		= New-Object System.Drawing.Point(10,58)
$UserIDLabel.Font                   	  		= 'Microsoft Sans Serif,8.25'

$UserIDTextbox                             		= New-Object system.Windows.Forms.TextBox
$UserIDTextbox.width                       		= 130
$UserIDTextbox.height                      		= 20
$UserIDTextbox.location                    		= New-Object System.Drawing.Point(70,55)
$UserIDTextbox.Font                       		= 'Microsoft Sans Serif,8.25'
#$UserIDTextbox.MaxLength = 15

$UserIDButton                         			= New-Object system.Windows.Forms.Button
$UserIDButton.text                    			= "Search"
$UserIDButton.width                   			= 60
$UserIDButton.height                  			= 20
$UserIDButton.location                			= New-Object System.Drawing.Point(210,55)
$UserIDButton.Font                    			= 'Microsoft Sans Serif,8.25'
$UserIDButton.TabStop = $False
$UserIDButton.Add_Click({ searchByUserID })
#$UserIDButton = createItem "Button" 210 55 60 20 "Search" $mainForm

#PC Name Info
$PCNameLabel                               		= New-Object system.Windows.Forms.Label
$PCNameLabel.text                           	= "PC Name: "
$PCNameLabel.width                          	= 60
$PCNameLabel.height                         	= 20
$PCNameLabel.location                       	= New-Object System.Drawing.Point(10,88)
$PCNameLabel.Font                           	= 'Microsoft Sans Serif,8.25'

$PCNameTextbox                       	      	= New-Object system.Windows.Forms.TextBox
$PCNameTextbox.width                          	= 130
$PCNameTextbox.height                         	= 20
$PCNameTextbox.location                       	= New-Object System.Drawing.Point(70,85)
$PCNameTextbox.Font                           	= 'Microsoft Sans Serif,8.25'
$PCNameTextbox.MaxLength = 15
#$PCNameTextbox = createItem "TextBox" 70 85 130 20 "" $mainForm

$PCSearchButton                             	= New-Object system.Windows.Forms.Button
$PCSearchButton.text                        	= "Search"
$PCSearchButton.width                       	= 60
$PCSearchButton.height                      	= 20
$PCSearchButton.location                    	= New-Object System.Drawing.Point(210,85)
$PCSearchButton.Font                        	= 'Microsoft Sans Serif,8.25'
$PCSearchButton.TabStop = $False
$PCSearchButton.Add_Click({ searchByPCName })

#$PCSearchButton = createItem "Button" 210 85 60 20 "Search" $mainForm
#IP Info
$IPAddressLabel                               	= New-Object system.Windows.Forms.Label
$IPAddressLabel.text                          	= "IP: "
$IPAddressLabel.width                         	= 18
$IPAddressLabel.height                       	= 20
$IPAddressLabel.location                      	= New-Object System.Drawing.Point(10,118)
$IPAddressLabel.Font                          	= 'Microsoft Sans Serif,8.25'

# IP Source Info
$IPAddressSourceLabel                        	= New-Object system.Windows.Forms.Label
$IPAddressSourceLabel.text                    	= "" # Starts Empty
$IPAddressSourceLabel.width                  	= 40
$IPAddressSourceLabel.height                  	= 20
$IPAddressSourceLabel.location                	= New-Object System.Drawing.Point(28,118)
$IPAddressSourceLabel.Font                    	= 'Microsoft Sans Serif,8.25'

$IPAddressTextbox                        	    = New-Object system.Windows.Forms.TextBox
$IPAddressTextbox.width                         = 130
$IPAddressTextbox.height                        = 20
$IPAddressTextbox.location                      = New-Object System.Drawing.Point(70,115)
$IPAddressTextbox.Font                          = 'Microsoft Sans Serif,8.25'
$IPAddressTextbox.MaxLength = 15
# $IPAddressTextbox = createItem "Textbox" 70 115 130 20 "" $mainForm
#	$IPAddressTextbox.MaxLength = 15

$IPAddressSearchButton                        	= New-Object system.Windows.Forms.Button
$IPAddressSearchButton.text                    	= "Search"
$IPAddressSearchButton.width                   	= 60
$IPAddressSearchButton.height                  	= 20
$IPAddressSearchButton.location                	= New-Object System.Drawing.Point(210,115)
$IPAddressSearchButton.Font                   	= 'Microsoft Sans Serif,8.25'
$IPAddressSearchButton.TabStop = $False  # Why is Tabstop set to $False
$IPAddressSearchButton.Add_Click({ searchByIP })
#Ph
#$IPAddressSearchButton = createItem "Button" 210 115 60 20 "Search" $mainForm
	
#PHone Info
$PhoneNumberLabel                            	= New-Object system.Windows.Forms.Label
$PhoneNumberLabel.text                       	= "Phone:"
$PhoneNumberLabel.width                      	= 60
$PhoneNumberLabel.height                     	= 20
$PhoneNumberLabel.location                   	= New-Object System.Drawing.Point(10,148)
$PhoneNumberLabel.Font                       	= 'Microsoft Sans Serif,8.25'

$PhoneNamberTextBox                        		= New-Object system.Windows.Forms.TextBox
$PhoneNamberTextBox.width                  		= 200
$PhoneNamberTextBox.height                 		= 20
$PhoneNamberTextBox.location               		= New-Object System.Drawing.Point(70,145)
$PhoneNamberTextBox.Font                   		= 'Microsoft Sans Serif,8.25'
$PhoneNamberTextBox.ReadOnly = $True
$PhoneNamberTextBox.TabStop = $False	# Why is Tabstop set to $False
#$PhoneNamberTextBox = createItem "Textbox" 70 145 200 20 "" $mainForm

#Lockout Info
$LockedOutUserLabel                         	= New-Object system.Windows.Forms.Label
$LockedOutUserLabel.text                     	= "Lockout:"
$LockedOutUserLabel.width                    	= 60
$LockedOutUserLabel.height                  	= 20
$LockedOutUserLabel.location                 	= New-Object System.Drawing.Point(10,178)
$LockedOutUserLabel.Font                     	= 'Microsoft Sans Serif,8.25'

$LockedOutUserTextbox                        	= New-Object system.Windows.Forms.TextBox
$LockedOutUserTextbox.width                  	= 130
$LockedOutUserTextbox.height                 	= 20
$LockedOutUserTextbox.location               	= New-Object System.Drawing.Point(70,175)
$LockedOutUserTextbox.Font                   	= 'Microsoft Sans Serif,8.25'
$LockedOutUserTextbox.ReadOnly 					= $True
$LockedOutUserTextbox.TabStop = $False # Why is Tabstop set to $False
 #$LockedOutUserTextbox = createItem "Textbox" 70 175 130 20 "" $mainForm

$LockedOutUserButton                         	= New-Object system.Windows.Forms.Button
$LockedOutUserButton.Text                   	= "Unlock"
$LockedOutUserButton.width                   	= 60
$LockedOutUserButton.height                  	= 20
$LockedOutUserButton.location                	= New-Object System.Drawing.Point(210,175)
$LockedOutUserButton.Font                    	= 'Microsoft Sans Serif,8.25'
$LockedOutUserButton.TabStop                 	= $False	# Why is Tabstop set to $False
$LockedOutUserButton.Visible                 	= $False # Hide button if useraccount is not locked out in AD.
$LockedOutUserButton.Add_Click({ unlockAccount })

#$LockedOutUserButton = createItem "Button" 210 175 60 20 "Unlock" $mainForm
	
#H Drive Info
$HDriveLabel                          			= New-Object system.Windows.Forms.Label
$HDriveLabel.text                     			= "H Drive:"
$HDriveLabel.width                    			= 60
$HDriveLabel.height                   			= 20
$HDriveLabel.location                 			= New-Object System.Drawing.Point(10,208)
$HDriveLabel.Font                    		 	= 'Microsoft Sans Serif,8.25'

$HDriveTextbox                        			= New-Object system.Windows.Forms.TextBox
$HDriveTextbox.width                  			= 200
$HDriveTextbox.height                 			= 20
$HDriveTextbox.location               			= New-Object System.Drawing.Point(70,208)
$HDriveTextbox.Font                   			= 'Microsoft Sans Serif,8.25'
$HDriveTextbox.ReadOnly                 		= $True
$HDriveTextbox.TabStop							= $False	# Why is Tabstop set to $False
#$HDriveTextbox = createItem "Textbox" 70 208 200 20 "" $mainForm

#OU Info
$OULabel                              			= New-Object system.Windows.Forms.Label
$OULabel.text                         			= "User OU:"
$OULabel.width                        			= 60
$OULabel.height                       			= 20
$OULabel.location                     			= New-Object System.Drawing.Point(10,238)
$OULabel.Font                         			= 'Microsoft Sans Serif,8.25'

$OUTextbox                              		= New-Object system.Windows.Forms.TextBox
$OUTextbox.width                       	 		= 200
$OUTextbox.height                       		= 20
$OUTextbox.location                     		= New-Object System.Drawing.Point(70,235)
$OUTextbox.Font                         		= 'Microsoft Sans Serif,8.25'
$OUTextbox.ReadOnly = $True
$OUTextbox.TabStop = $False # Why is Tabstop set to $False
#$OUTextbox = createItem "Textbox" 70 235 200 20 "" $mainForm
#Buttons

$SCCMRemoteControlButton                      	= New-Object system.Windows.Forms.Button
$SCCMRemoteControlButton.text                  	= "Connect Via SCCM Remote Control"
$SCCMRemoteControlButton.width                	= 259
$SCCMRemoteControlButton.height               	= 20
$SCCMRemoteControlButton.location              	= New-Object System.Drawing.Point(10,265)
$SCCMRemoteControlButton.Font                  	= 'Microsoft Sans Serif,8.25'
$SCCMRemoteControlButton.Add_Click({ runRemoteViewer })

# $SCCMRemoteControlButton = createItem "Button" 10 265 259 20 "Connect Via SCCM Remote Control" $mainForm
$UserDetailsButton                         		= New-Object system.Windows.Forms.Button
$UserDetailsButton.text                   		= "User Details"
$UserDetailsButton.width                   		= 122
$UserDetailsButton.height                  		= 20
$UserDetailsButton.location                		= New-Object System.Drawing.Point(10,295)
$UserDetailsButton.Font                    		= 'Microsoft Sans Serif,8.25'
$UserDetailsButton.Add_Click({ runUserFacts })
# $UserDetailsButton = createItem "Button" 10 295 122 20 "User Details" $mainForm

$UserOUGroupsButton                         	= New-Object system.Windows.Forms.Button
$UserOUGroupsButton.text                    	= "User OU Groups"
$UserOUGroupsButton.width                   	= 122
$UserOUGroupsButton.height                  	= 20
$UserOUGroupsButton.location                	= New-Object System.Drawing.Point(10,325)
$UserOUGroupsButton.Font                   	 	= 'Microsoft Sans Serif,8.25'
$UserOUGroupsButton.Add_Click({ runUserGroups })
#$UserOUGroupsButton = createItem "Button" 10 325 122 20 "User Groups" $mainForm
$ChangePasswordButton                         	= New-Object system.Windows.Forms.Button
$ChangePasswordButton.text                    	= "Change Password"
$ChangePasswordButton.width                   	= 122
$ChangePasswordButton.height                  	= 20
$ChangePasswordButton.location                	= New-Object System.Drawing.Point(10,355)
$ChangePasswordButton.Font                    	= 'Microsoft Sans Serif,8.25'
$ChangePasswordButton.Add_Click({newUserPassword})
# $ChangePasswordButton = createItem "Button" 10 355 122 20 "Change Password" $mainForm
$PCDetailsButton                         		= New-Object system.Windows.Forms.Button
$PCDetailsButton.text                    		= "PC Details"
$PCDetailsButton.width                   		= 122
$PCDetailsButton.height                  		= 20
$PCDetailsButton.location                		= New-Object System.Drawing.Point(147,295)
$PCDetailsButton.Font                    		= 'Microsoft Sans Serif,8.25'
$PCDetailsButton.Add_Click({ runPCFacts })
# $PCDetailsButton = createItem "Button" 147 295 122 20 "PC Details" $mainForm
$PCOUGroups                         			= New-Object system.Windows.Forms.Button
$PCOUGroups.text                    			= "PC OU Groups"
$PCOUGroups.width                   			= 122
$PCOUGroups.height                  			= 20
$PCOUGroups.location                			= New-Object System.Drawing.Point(147,325)
$PCOUGroups.Font                    			= 'Microsoft Sans Serif,8.25'
$PCOUGroups.Add_Click({ runPCGroups })	
#$PCOUGroups = createItem "Button" 147 325 122 20 "PC Groups" $mainForm
$ComputerManagementButton                     	= New-Object system.Windows.Forms.Button
$ComputerManagementButton.text                 	= "Computer Mgt"
$ComputerManagementButton.width                	= 122
$ComputerManagementButton.height               	= 20
$ComputerManagementButton.location             	= New-Object System.Drawing.Point(147,355)
$ComputerManagementButton.Font                 	= 'Microsoft Sans Serif,8.25'
$ComputerManagementButton.Add_Click({ runManagePC })
#$ComputerManagementButton = createItem "Button" 147 355 122 20 "Manage PC" $mainForm
$ViewRemotePCCDrive                         	= New-Object system.Windows.Forms.Button
$ViewRemotePCCDrive.text                    	= "View Remote C:"
$ViewRemotePCCDrive.width                   	= 122
$ViewRemotePCCDrive.height                  	= 20
$ViewRemotePCCDrive.location                	= New-Object System.Drawing.Point(147,385)
$ViewRemotePCCDrive.Font                    	= 'Microsoft Sans Serif,8.25'
$ViewRemotePCCDrive.Add_Click({ runViewC })

# Use Remote Desktop to connect to remote machine
$RDPButton                         				= New-Object system.Windows.Forms.Button
$RDPButton.text                    				= "RDP"
$RDPButton.width                   				= 122
$RDPButton.height                  				= 20
$RDPButton.location                				= New-Object System.Drawing.Point(147,415)
$RDPButton.Font                    				= 'Microsoft Sans Serif,8.25'
$RDPButton.Add_Click({ runRDP })

# Use PowerShell Session to connect to remote machine
$PowerShellRemoteButton                         = New-Object system.Windows.Forms.Button
$PowerShellRemoteButton.text                    = "PS Remote"
$PowerShellRemoteButton.width                   = 122
$PowerShellRemoteButton.height                  = 20
$PowerShellRemoteButton.location                = New-Object System.Drawing.Point(147,445)
$PowerShellRemoteButton.Font                    = 'Microsoft Sans Serif,8.25'	
$PowerShellRemoteButton.Add_Click({connectPSSession})

# Make Wrench stay on top of all windows
$KeepWrenchTopMostCheckbox                    	= New-Object system.Windows.Forms.CheckBox
$KeepWrenchTopMostCheckbox.text              	= "Keep Wrench on Top"
$KeepWrenchTopMostCheckbox.AutoSize           	= $false
$KeepWrenchTopMostCheckbox.width               	= 100
$KeepWrenchTopMostCheckbox.height              	= 35
$KeepWrenchTopMostCheckbox.location            	= New-Object System.Drawing.Point(30,410)
$KeepWrenchTopMostCheckbox.Font                	= 'Microsoft Sans Serif,8.25'
$KeepWrenchTopMostCheckbox.Add_Click({ runOnTop })
#$KeepWrenchTopMostCheckbox = createItem "Checkbox" 30 410 100 35 "Keep Wrench on Top" $mainForm

$NewLocalPowerShellWindowLabel              	= New-Object system.Windows.Forms.Label
$NewLocalPowerShellWindowLabel.text           	= "New Local Powershell Window"
$NewLocalPowerShellWindowLabel.width          	= 150
$NewLocalPowerShellWindowLabel.height           = 15
$NewLocalPowerShellWindowLabel.location       	= New-Object System.Drawing.Point(80,578)
$NewLocalPowerShellWindowLabel.Font             = 'Microsoft Sans Serif,8.25'
$NewLocalPowerShellWindowLabel.ForeColor 		= "Blue"
$NewLocalPowerShellWindowLabel.Add_Click({ Start-Process "powershell.exe" })

$PingTimer = New-Object System.Windows.Forms.Timer
$PingTimer.Interval = 1000
$PingTimer.add_tick({ checkPing })

# Wrench Logo PictureBox Image
$WrenchLogoPictureBox = new-object Windows.Forms.PictureBox
$WrenchLogoPictureBox.location = New-Object System.Drawing.Size(10,477)
$WrenchLogoPictureBox.size = New-Object System.Drawing.Size(260,100)
$WrenchLogoPictureBox.BorderStyle = "FixedSingle"
$WrenchLogoPictureBox.Image = [System.Drawing.Image]::Fromfile((get-item $LogoLocation));

#Expand Button
$ExpandButton                        			= New-Object system.Windows.Forms.Button
$ExpandButton.text                    			= ">"
$ExpandButton.width                   			= 15
$ExpandButton.height                  			= 15
$ExpandButton.location                			= New-Object System.Drawing.Point(264,577)
$ExpandButton.Font                    			= 'Microsoft Sans Serif,8.25'
$ExpandButton.Visible 				 			= $true #this is not needed.  The default for a button is visible
$ExpandButton.Add_Click({ expandForm })

# Add Labels to MainForm
$mainForm.controls.AddRange(@($UserNameLabel,$UserIDLabel,$IPAddressLabel,
$PCNameLabel,$IPAddressSourceLabel,$PhoneNumberLabel,$LockedOutUserLabel,
$HDriveLabel,$OULabel,$NewLocalPowerShellWindowLabel))
# Add Textboxes to MainForm 
$mainForm.controls.AddRange(@($UserNameTextbox,$IPAddressTextbox,
$PCNameTextbox,$OUTextbox,$UserIDTextbox,$PhoneNumberLabel,$PhoneNamberTextBox,
$HDriveTextbox,$LockedOutUserTextbox)) 
# Add Buttons to MainForm
$mainForm.controls.AddRange(@($UserNameSearchButton,$UserIDButton,
$IPAddressSearchButton,$LockedOutUserButton,$SCCMRemoteControlButton,
$UserDetailsButton,$UserOUGroupsButton,$ChangePasswordButton,$PowerShellRemoteButton,
$RDPButton,$ExpandButton,$PCSearchButton,$PCDetailsButton,$PCOUGroups,
$ComputerManagementButton,$RenameButton,$SCCMClientCenterButton,$ViewRemotePCCDrive,
$ViewRemotePCCDrive,$GPBtn,$TelnetPCButton,$RenameButton,$MSRemoteAssistanceButton))
# Add Other Controls that are not defined above
$mainForm.controls.AddRange(@($KeepWrenchTopMostCheckbox,$WrenchLogoPictureBox))
####### EXPANDED FORM GUI #######

#Draw Seperator
$pen = new-object Drawing.SolidBrush black
$formGraphics = $mainForm.createGraphics()
$mainForm.add_paint({$formGraphics.DrawLine($pen, 279, 1, 279, 620)})
	
#Buttons
$GPBtn                         = New-Object system.Windows.Forms.Button
$GPBtn.text                    = "Check Group Policy"
$GPBtn.width                   = 122
$GPBtn.height                  = 20
$GPBtn.location                = New-Object System.Drawing.Point(290,25)
$GPBtn.Font                    = 'Microsoft Sans Serif,8.25'
$GPBtn.Add_Click({ pullGroupPolicy })
#$GPBtn = createItem "Button" 290 25 122 20 "Check Group Policy" $mainForm
$GPTimer = New-Object System.Windows.Forms.Timer
$GPTimer.Interval = 2000
$GPTimer.add_tick({ checkGP })

$TelnetPCButton                         		= New-Object system.Windows.Forms.Button
$TelnetPCButton.text                    		= "Telnet"
$TelnetPCButton.width                  			= 122
$TelnetPCButton.height                  		= 20
$TelnetPCButton.location                		= New-Object System.Drawing.Point(290,55)
$TelnetPCButton.Font                    		= 'Microsoft Sans Serif,8.25'
$TelnetPCButton.Add_Click({ runTelnet })

#$TelnetPCButton = createItem "Button" 290 55 122 20 "Telnet" $mainForm
$RenameButton                         			= New-Object system.Windows.Forms.Button
$RenameButton.text                    			= "Rename PC"
$RenameButton.width                   			= 122
$RenameButton.height                  			= 20
$RenameButton.location                			= New-Object System.Drawing.Point(290,85)
$RenameButton.Font                    			= 'Microsoft Sans Serif,8.25'
$RenameButton.Add_Click({ runRename })
# $RenameButton = createItem "Button" 290 85 122 20 "Rename PC" $mainForm
$SCCMClientCenterButton                         = New-Object system.Windows.Forms.Button
$SCCMClientCenterButton.text                    = "Client Center"
$SCCMClientCenterButton.width                   = 122
$SCCMClientCenterButton.height                  = 20
$SCCMClientCenterButton.location                = New-Object System.Drawing.Point(290,115)
$SCCMClientCenterButton.Font                    = 'Microsoft Sans Serif,8.25'
$SCCMClientCenterButton.Add_Click({ openClientCenter })	
# $SCCMClientCenterButton = createItem "Button" 290 115 122 20 "Client Center" $mainForm
$MSRemoteAssistanceButton                     	= New-Object system.Windows.Forms.Button
$MSRemoteAssistanceButton.text                 	= "MS Remote Assist"
$MSRemoteAssistanceButton.width                	= 122
$MSRemoteAssistanceButton.height              	= 20
$MSRemoteAssistanceButton.location              = New-Object System.Drawing.Point(290,145)
$MSRemoteAssistanceButton.Font                 	= 'Microsoft Sans Serif,8.25'
$MSRemoteAssistanceButton.Add_Click({ openMSRA })
#$MSRemoteAssistanceButton = createItem "Button" 290 145 122 20 "MS Remote Assist" $mainForm

#Labels
#$CMDriveLbl = createItem "Label" 290 145 122 20 "SCCM Site: $CMDrive" $mainForm
#$PSScriptRootlbl = createItem "Label" 290 160 122 60 "PSScriptRoot: $PSScriptRoot" $mainForm

####### VARIABLES #######
$global:Name = ""
$global:UserID = ""
$global:PCName = ""
$global:IP = ""
$global:Lockout = ""
$global:Phone = ""
$global:ValidName = $False
$global:ValidUserID = $False
$global:ValidPCName = $False
$global:IPWithARec = $True



### MAIN FORM FUNCTIONS ###
function searchByName{
	clearVariables
	$global:Name = ($UserNameTextbox.text).Trim()
	testName		
	if ($global:validName -eq $True){
		if(!(pickName)){return}
		clearBoxes
		$UserNameTextbox.text = $global:Name
		$LockedOutUserButton.Visible=$false
		if(!(IDByLastName)){return}
		PhoneByUserID
		if(!(PCNameByUserID)){return}
		LockoutByUserID
		HDriveByUserID
		OUByUserID
		getIP
		pingIP
	}else{
		$msg::Show("No User Found")
	}
}
function searchByUserID{
	clearVariables
	$global:UserID = ($UserIDTextbox.Text).Trim() #FIXTHIS
	testID
	if($global:validID -eq $True){
		if(!(pickID)){return}
		clearBoxes
		$UserIDTextbox.Text = $global:UserID
		$LockedOutUserButton.Visible=$false
		NameByUserID
		if(!(PCNameByUserID)){return}
		LockoutByUserID
		PhoneByUserID
		HDriveByUserID
		OUByUserID
		getIP
		pingIP
	}else{
		$msg::Show("No UserID found")
	}
}
$global:UserID
function searchByPCName{
	clearVariables
	$global:PCName = ($PCNameTextbox.Text).Trim()
	testPCName
	if($global:validPCName -eq $True){
		if(!(pickPCName)){return}
		clearBoxes
		$PCNameTextbox.Text = $global:PCName
		$LockedOutUserButton.Visible=$false
		UserIDByPCName
		NameByUserID
		LockoutByUserID
		PhoneByUserID
		HDriveByUserID
		OUByUserID
		getIP
		pingIP
	}else{
		$msg::Show("No PC found")
	}
}
function searchByIP{
	clearVariables
	$global:IP = ($IPAddressTextbox.Text).Trim()
	if(testIP){
		clearBoxes
		$IPAddressTextbox.Text = $global:IP
		if($global:IPWithARec){
			$LockedOutUserButton.Visible=$false
			PCNameByIP
			UserIDByPCName
			NameByUserID
			LockoutByUserID
			PhoneByUserID
			HDriveByUserID
			OUByUserID
		}
	}

}
function unlockAccount{
	try{
		Unlock-ADAccount $global:UserID -Confirm:$False
		if((Get-ADUser $global:UserID -properties LockedOut).LockedOut -eq $False){
			$LockedOutUserButton.Visible = $False
			$LockedOutUserTextbox.Text = "Not Locked"
		}else{
			$LockedOutUserTextbox.Text = "Locked"
		}
	}catch{
		$msg::Show($error[0].toString() + "`r`n`r`nIf this is in error, please check out http://support.microsoft.com/kb/2577917")
	}
}
function runRemoteViewer{
	if ($global:IP -eq "" -or $global:IP -eq "-"){
		$msg::Show("Search PC or IP first!")
	}else{
		& $SCCMRemoteLocation $global:IP
	}
}
function runUserFacts{
	if($global:userID -eq "" -or $global:userID -eq "-"){
		$msg::Show("Find a user first!")
	}else{
		extraUserFacts
	}
}
function runUserGroups{
	if($global:userID -eq "" -or $global:userID -eq "-"){
		$msg::Show("Find a user first!")
	}else{
		viewGroups "user"
	}
}
function runPCFacts{
	if($global:PCName -eq "" -or $global:PCName -eq "-"){
		$msg::Show("Find a PC first!")
	}else{
		extraPCFacts
	}
}
function runPCGroups{
	if($global:PCName -eq "" -or $global:PCName -eq "-"){
		$msg::Show("Find a PC first!")
	}else{
		viewGroups "computer"
	}
}
function runManagePC{
	if ($global:pcname -eq "-" -and $global:IP -eq "-"){
		$msg::Show("Pick a PC first")
	}else{
		try{
			compmgmt.msc /computer:$global:pcname
		}catch{
			$msg::Show($error[0])
		}
	}
}
function runViewC{
	if ($global:pcname -eq "-" -and $global:IP -eq "-"){
		$msg::Show("Pick a PC first")
	}else{
		try{
			viewCFolder
		}catch{
			$msg::Show($error[0])
		}
	}
}
function runRDP{
	try{
		Start-Process "mstsc.exe" "/v:$global:PCName"
	}catch{
		$msg::Show($error[0])
	}
}
function connectPSSession{
	$commandArg = '-Command "Enter-PSSession -ComputerName ' + $PCName
	Start-Process "powershell.exe" -ArgumentList '-NoExit', $commandArg
}
function runOnTop{
	if ($KeepWrenchTopMostCheckbox.Checked -eq $true){
		$mainForm.TopMost = $true
		$mainForm.Update()
	}else{
		$mainForm.TopMost = $false
		$mainForm.Update()
	}
}
function mainKeyboard{
	if ($_.KeyCode -eq "Enter"){
		if($UserNameTextbox.Focused){ $UserNameSearchButton.PerformClick() }
		elseif ($PCNameTextbox.Focused){ $PCSearchButton.PerformClick() }
		elseif ($IPAddressTextbox.Focused){ $IPAddressSearchButton.PerformClick() }
		elseif ($UserIDTextbox.Focused){ $UserIDButton.PerformClick() }
	}
}

### EXPANDED FORM FUNCTIONS ###
function pullGroupPolicy{
	$resultpath = $env:temp + "\GPResult.html"
	remove-item $resultpath
	if($global:PCName -eq ""){
		$msg::Show("Please find a computer first")
	}elseif($global:PCName -ne "" -and $global:userID -eq ""){
		#Runs if PCName exists but username doesn't
		$global:GPJob = Start-Job {
			param($PCName,$resultpath)
			try{
				gpresult /s $global:PCName /scope:computer /h $resultpath
			}catch{	$msg::Show($error[0]) }
		} -Arg @($global:PCName,$resultpath)

	} elseif($global:PCName -ne "" -and $global:userID -ne ""){
		#Runs if PCName and UserID are there
		$global:GPJob = Start-Job {
			param($PCName, $userID, $resultpath)
			try{
				gpresult /s $global:PCName /user $global:UserID /h $resultpath
			}catch{	$msg::Show($error[0]) }	
		} -Arg @($global:PCName,$global:userID,$resultpath)
	}
	$GPTimer.Enabled = $true
	$GPBtn.Forecolor = "blue"
}
function checkGP{
	$state = $global:GPJob.State
	$resultpath = $env:temp + "\GPResult.html"
	if (($global:GPJob.State -eq "Completed")){
		if (Test-Path $resultPath){
			try{
				$IE = New-Object -com internetexplorer.application
				$IE.visible = $true
				$IE.navigate($resultpath) 
			}catch{ $msg::Show($error[0]) }
			}
		$GPTimer.Enabled = $false
		$GPBtn.Forecolor = "black"
	}
}
function runTelnet{
	if ($global:pcname -eq "-" -and $global:IP -eq "-"){
		$msg::Show("Pick a PC first")
	}else{
		try{
			telnetPC
		}catch{
			$msg::Show($error[0])
		}
	}
}
function runRename{
	try{
		$ProcessName = "/c start cmd /c cscript.exe " + $RenameComputerLocation
		Start-Process "cmd.exe" $ProcessName
	}catch{
		$msg::Show($error[0])
	}
}
function openClientCenter{
	& $ClientCenterLocation $global:pcname
}
function openMSRA {
	& msra /offerra $global:IP
}


### MAIN FORM UTILITY FUNCTIONS ###
function makeClickList([ref]$variable, $array, [ref]$returnval){
	function EnterListener(){
		if ($_.KeyCode -eq "Enter"){
			$variable.value = $pickList.SelectedItem
			$true
			$pickForm.close()
		}elseif ($_.Keycode -eq "Escape"){
			$false
			$pickForm.close()
		}
	}

	$pickForm = createForm "Select One:" 200 180 "CenterScreen" "Fixed3D" $false $false $false
		$pickForm.KeyPreview = $True	#Makes form aware of key presses so it sees "ENTER"
		$pickForm.Add_KeyDown({$returnval.value = EnterListener})
	$pickList = createItem "ListBox" 17 17 150 100 "" $pickForm
		$pickList.DataSource = $array
		$pickList.add_MouseDoubleClick({$variable.value = $pickList.SelectedItem;$returnval.value=$true;$pickForm.close()})
	showForm($PickForm)
}
function clearBoxes{
	$UserNameTextbox.Text = ""
	$UserIDTextbox.Text = ""
	$PCNameTextbox.Text = ""
	$IPAddressTextbox.Text = ""
	$LockedOutUserTextbox.Text = ""
	$PhoneNamberTextBox.Text = ""
}
function clearVariables{
	$global:Name = ""
	$global:UserID = ""
	$global:PCName = ""
	$global:IP = ""
	$global:Lockout = ""
	$global:Phone = ""
	$global:ValidName = $False
	$global:ValidUserID = $False
	$global:ValidPCName = $False
	$global:IPWithARec = $True
	$IPAddressTextbox.Forecolor = "black"
}
function subWindowKeyListener($form){
	if($_.Control -eq $true -and $_.KeyCode -eq "W" ){
		$form.close()
	}
}

### USER SEARCH FUNCTIONS
function pickName{
	$spaceindex = $name.IndexOf(" ")
	$commaindex = $name.IndexOf(",")
	$matches = @()
	if ($commaindex -gt 0){
		$search = $name + "*"
		$matches = @((Get-ADUser -f {name -like $search}).name)
	}else{
		if($spaceindex -gt 0){
			$firstname = $name.substring(0, $spaceindex)
			$lastname = $name.substring($spaceindex)
			$firstname += "*"
			$firstnamematches = @((Get-ADUser -f {GivenName -like $firstname}).name)
			$lastname+= "*"
			$lastnamematches = @((Get-ADUser -f {Surname -like $lastname}).name)
			ForEach ($fname in $firstnamematches){
				ForEach ($lname in $lastnamematches){
					if($fname -eq $lname){
						$matches += $fname
					}
				}
			}
		}else{
			$search = $name + "*"
			$firstnamematches = @((Get-ADUser -f {GivenName -like $search}).name)
			$lastnamematches = @((Get-ADUser -f {Surname -like $search}).name)
			if (($firstnamematches | Measure-Object).count -gt 0){
				$matches += $firstnamematches
			}
			if (($lastnamematches | Measure-Object).count -gt 0){
				$matches += $lastnamematches
			}
		}
	}
	
	if ($matches.count -eq 1){		
		$global:Name = $matches[0]
		$UserNameTextbox.Text = $global:Name
		$true
	}elseif ($matches.count -gt 1){
		$returnval = $false
		MakeClickList ([ref]$global:Name) $matches ([ref]$returnval)
		if(!$returnval){$false}else{$true}
	}
}
function pickID{
	$variantname = "$UserID" + "*"
	$ids = @((get-aduser -f {SamAccountName -like $variantname}).SamAccountName)
	if ($ids.length -eq 1){
		$global:UserID = $ids[0]
		$true
	}else{
		$returnval = $false
		MakeClickList ([ref]$global:UserID) $ids ([ref]$returnval)
		if(!$returnval){$false}else{$true}
	}
	$UserIDTextbox.Text = $global:UserID
}
function pickPCName{
	$variantname = "$PCName" + "*"
	$pcs = @((get-adcomputer -f {name -like $variantname}).name)
	if ($pcs.length -eq 1){
		$global:PCName = $pcs[0]
		$true
	}else{
		$returnval = $false
		MakeClickList ([ref]$global:PCName) $pcs ([ref]$returnval)
		if(!$returnval){$false}else{$true}
		
	}
	$PCNameTextbox.Text = $global:PCName
}
function testName{
	if (($Name).Length -lt 1){
		$global:validName = $False
	}else{
		$spaceindex = $name.IndexOf(" ")
		$commaindex = $name.IndexOf(",")
		$matches = @()
		if ($commaindex -gt 0){
			$search = $name + "*"
			$matches = Get-ADUser -f {name -like $search}
		}else{
			if($spaceindex -gt 0){
				$firstname = $name.substring(0, $spaceindex)
				$lastname = $name.substring($spaceindex)
				$firstname += "*"
				$firstnamematches = @(Get-ADUser -f {GivenName -like $firstname})
				$lastname+= "*"
				$lastnamematches = @(Get-ADUser -f {Surname -like $lastname})
				ForEach ($fname in $firstnamematches){
					ForEach ($lname in $lastnamematches){
						if($fname.sid -eq $lname.sid){
							$matches += $fname
						}
					}
				}
			}else{
				$search = $name + "*"
				$firstnamematches = @(Get-ADUser -f {GivenName -like $search})
				$lastnamematches = @(Get-ADUser -f {Surname -like $search})
				$matches = $firstnamematches + $lastnamematches
			}
		}
	}
	if (($matches | Measure-Object).count -lt 1){
		$global:validName = $false
	}else{
		$global:ValidName = $true
	}
}
function testID{
	if (($UserIDTextbox.Text).length -lt 1){
		$global:validID = $False
	}else{
		#if($UserIDTextbox.Text -contains "@"){
		$variantname = "$UserID" + "*"
		#$ids = @((get-aduser -f {UserPrincipalName -like $variantname}).UserPrincipalName)
		#}else{
		#	variantname = "$UserID" + "*"
			$ids = @((get-aduser -f {SamAccountName -like $variantname}).SamAccountName)
			if ($ids[0].length -lt 1){
				$global:validID = $False
			}else{
				$global:validID = $True
			}
		#}
	}
}
function testPCName{
	if (($PCNameTextbox.Text).length -lt 1){
		$global:validPCName = $False
	}else{
		$variantname = "$PCName" + "*"
		$pcs = @((get-adcomputer -f {name -like $variantname}).name)
		if ($pcs[0].length -lt 1){
			$global:validPCName = $False
		}else{
			$global:validPCName = $True
		}
	}

}
function testIP{
	$Online = Test-Connection -Computername $global:IP -BufferSize 16 -Count 1 -quiet
	if (!$Online){
		$false
		$IPAddressTextbox.Forecolor = "red"
	}else{
		try{
			$IPObject = [System.Net.IPAddress]::parse($IP)
			[System.Net.IPAddress]::tryparse([string]$IP, [ref]$IPObject)
			([System.Net.Dns]::GetHostByAddress($global:IP)).HostName
		}catch{
			$global:IPWithARec = $False
		}
		$IPAddressTextbox.Forecolor = "green"
		$true
	}

}
function IDByLastName{
	$variantname = "$Name" + "*"
	$ID = @((get-aduser -f {name -like $variantname}).SamAccountName)
	if($ID.length -eq 1){
		$global:UserID = $ID[0]
		$true
	}else{
		$returnval = $false
		MakeClickList ([ref]$global:UserID) $ID ([ref]$returnval)
		if(!$returnval){$false}else{$true}
		
	}
	
	$UserIDTextbox.Text = $global:UserID
}

function PCNameByUserID{
	
	# Check if $SCCMSiteCode is undefined. Get Site Code from local machine
	If (([string]::IsNullOrWhiteSpace($($SCCMSiteCode)))){
		$SCCMSiteCode = ([wmiclass]"ROOT\ccm:SMS_Client").GetAssignedSite().sSiteCode
		If (([string]::IsNullOrWhiteSpace($($SCCMSiteCode)))) {
			[System.Windows.MessageBox]::Show("Missing Variable $SCCMSiteServer`nAdd The name of the SCCM Server to Wrench_env.ps1")
		}
	}
	# Check if $SCCMSiteServer is undefined. Get SCCM Site Server from local machine
	If (([string]::IsNullOrWhiteSpace($($SCCMSiteServer)))){
		$SCCMSiteServer =  ([wmi]"ROOT\ccm:SMS_Authority.Name='SMS:$SCCMSiteCode'").CurrentManagementPoint # If the SCCMSiteServer is not in config file get it for the local machine.
		If (([string]::IsNullOrWhiteSpace($($SCCMSiteServer)))) {
			[System.Windows.MessageBox]::Show('Missing Variable $SCCMSiteServer`nAdd The name of the SCCM Server to Wrench_env.ps1')
			# future update give the user a textbox to update $SCCMSiteServer
		}
		
	}
	$variantid = "$UserID" + "*"
	$pcnames = @((Get-WmiObject -Class SMS_UserMachineRelationship -namespace "root\sms\site_$SCCMSiteCode" -computer $SCCMSiteServer -filter "UniqueUserName LIKE '%$global:UserID' and Types='1' and IsActive='1' and Sources='4'").ResourceName)
			
	if ($pcnames.length -eq 1){
		$global:PCName = $pcnames[0]
		$true
	}else{
		$returnval = $false
		MakeClickList ([ref]$global:PCName) $pcnames ([ref]$returnval)
		if(!$returnval){$false}else{$true}
	}
	if ($global:PCName.length -lt 2){
		$global:PCName = "-"
		$PCNameTextbox.Text = "-"
	}
		$PCNameTextbox.Text = $global:PCName
}

function getIP{
	if($global:PCName -eq "-"){
		$global:IP = "-"
	}else{
		if($global:Name -ne "" -and $global:name -ne "-"){getVPNIP}
		if(!($global:IP -like "*.*.*.*")){
			getADIP
			if(!($global:IP -like "*.*.*.*")){
				if(!(getDNSIP)){return}
			}
		}
	}
	$IPAddressTextbox.Text = $global:IP
}
function LockoutByUserID{
	if ($global:UserID -ne "-"){
		$user = get-aduser $global:userID -properties LockedOut
		if ($user.LockedOut -eq $True){
			$LockedOutUserTextbox.Text = "Locked"
			$LockedOutUserButton.Visible = $True

		}else{
			$LockedOutUserTextbox.Text = "Not Locked"
		}
	}else{
		$global:Lockout = "-"
		$LockedOutUserTextbox.Text = "-"
	}
		
}
function PhoneByUserID{
			if ($global:UserID -ne "-"){
				$UserPhoneNumber = Get-ADUser  $global:UserID -properties Pager, ipPhone, OfficePhone, MobilePhone
				If (-not([string]::IsNullOrWhiteSpace(($UserPhoneNumber).Pager)))  # If the value is NOT $null, Empty string "" , or any number of spaces "       " this equals true
				{ 
					$global:Phone = $UserPhoneNumber.Pager
				}
				Elseif (-not([string]::IsNullOrWhiteSpace($UserPhoneNumber.ipPhone)))
				{
					$global:Phone = $UserPhoneNumber.ipPhone
				} 
				Elseif (-not([string]::IsNullOrWhiteSpace($UserPhoneNumber.OfficePhone)))
				{
					$global:Phone = $UserPhoneNumber.OfficePhone
				} 
				Elseif (-not([string]::IsNullOrWhiteSpace($UserPhoneNumber.MobilePhone)))
				{
					$global:Phone = $UserPhoneNumber.MobilePhone
				} 
				Else {$global:Phone = '-'} # There is no phone number in AD. 		
			}
			Else {$global:Phone = '-'}  # The UserName field has a "-" in it.
	
			$PhoneNamberTextBox.Text = $global:Phone
		}

function NameByUserID{
	if ($global:UserID -ne "-"){
		$global:Name = (Get-ADUser $global:UserID).Name
	}else{
		$global:Name = "-"
	}	

	$UserNameTextbox.Text = $global:Name
}
function UserIDByPCName{
	if($global:PCName.length -lt 7){
		$global:UserID = "-"
	}else{
		$variantID = ($global:PCName).Substring(0,7)
		$variantname = "$variantID" + "*"
		$UserName = @(((Get-WmiObject -Class SMS_UserMachineRelationship -namespace "root\sms\$SCCMNameSpace" -computer $SCCMSiteServer -filter "ResourceName='$global:PCName' and Types='1' and IsActive='1' and Sources='4'").UniqueUserName).substring($Domain.length+1))
		$SamAccounts=foreach ($name in $UserName){@((get-aduser -f {SamAccountName -like $name}).SamAccountName)}
		$UPN= foreach ($Account in $SamAccounts){@((get-aduser -f {SamAccountName -like $name}).UserPrincipalName)}
		
		if ($UserName.length -eq 1){
			$global:UserID = $UserName[0]
			$true
		}elseif($UserName.length -eq 0 ){
			$global:UserID = "-"
			$true
		}else{
			$returnval = $false
			MakeClickList ([ref]$global:UserID) $UserName ([ref]$returnval)
			if(!$returnval){$false}else{$true}
			
		}
	}
	$UserIDTextbox.Text = $global:UserID	
}
function PCNameByIP{
	$hostname = ([System.Net.Dns]::GetHostByAddress($IP)).HostName
	if ($hostname.indexof('.') -eq -1){
		$global:PCName = $hostname
	}else{
		$global:PCName = $hostname.Substring(0,$hostname.IndexOf('.'))
	}
	$PCNameTextbox.Text = $PCName

	
}
function HDriveByUserID{
	try{
		if ($global:UserID -ne "-"){
			$user = Get-AdUser $global:UserID -Properties HomeDirectory
			$HDriveTextbox.Text = $user.HomeDirectory
		}else{
			$HDriveTextbox.Text = "-"
		}
	}catch{
	
	}
	
}
function OUByUserID{
	try{
		if ($global:UserID -ne "-"){
			$user = Get-AdUser $global:UserID -Properties CanonicalName
			$OUTextbox.Text = $user.CanonicalName.Substring(($user.CanonicalName).IndexOf('/'))
		}else{
			$OUTextbox.Text = "-"
		}
	}catch{
	
	}
	
}

### IP FUNCTIONS
function getVPNList{
	$env = get-content -Path ($env:USERPROFILE + "\vpnenv.txt")
	try{	
		$username = $env[0]
		$password = $env[1]
		
		[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true} #Ignore SSL cert
		$webclient = new-object System.Net.WebClient
		$webclient.Credentials = new-object System.Net.NetworkCredential($username, $password)
		$webpage = $webclient.DownloadString($VPNSiteUrl)	# Ateempt to access website with credentials
	}catch{
		$msg::Show($error[0]) 
	}
	$webpage
}
function getNameIPArray($split){
	$vpnarray = @()
	for ($i = 0; $i -lt $split.length; $i++){
		$line = $split[$i]
		$nextline = $split[$i+1]
		
		if ($line.indexof(',') -gt 0 -and $nextline -like "*.*.*.*"){
			$vpnip = ($nextline.split(" "))[4]
			$name = ($line.split(" "))[6] + " " + ($line.split(" "))[7]
			$vpnarray += @(,($name,$vpnip))
		}elseif($nextline -like "*.*.*.*"){
			$name = ($line.split(" "))[6]
			$vpnip = ($nextline.split(" "))[4]
			$vpnarray += @(,($name,$vpnip))
		}
	}
	$vpnarray
}
function searchVPNArray($search, $array){
	if (!($search -like "*.*.*.*") -and $name.indexof(',') -ne -1){
		$firstname = $name.Substring($name.indexof(" ")+1)
		$lastname = $name.Substring(0, $name.indexof(','))
		#Name Types
		$lastcommafirst = "*" + $name + "*"
		$firstdotlast = "*" + $firstname + "." + $lastname + "*"
		$lastdotfirst = "*" + $lastname + "." + $firstname + "*"
		$lastf = "*" + $lastname + $firstname[0] + "*"
	}
	ForEach ($pair in $array){
		if ($pair[1] -like $search){
			$pair
		}elseif ($pair[0] -like $lastcommafirst -or $pair[0] -like $firstdotlast -or $pair[0] -like $lastdotfirst -or $pair[0] -like $lastf){
			$pair
		}
	}
}
function getVPNIP{
	$env = get-content -Path ($env:USERPROFILE + "\vpnenv.txt")
	if ($env.count -gt 1){
		$webpage = getVPNList
		$split = ($webpage -split '[\r\n]') | Where-Object {$_}  #Split page into lines
		if($split.count -gt 2){
			$NameIPArray = getNameIPArray $split
			$pair = searchVPNArray $global:name $NameIPArray
			if ($pair.count -eq 2){
				$global:IP = $pair[1]
				$IPAddressSourceLabel.Text ="(vpn)"
			}
		}
	}
}
function getADIP{
	try{
		$global:IP = (Get-ADComputer $global:PCName -Properties IPv4Address).IPv4Address
		$IPAddressSourceLabel.Text ="(ad)"
	}catch{
		$global:IP = "-"
	}
}
function getDNSIP{
	$ips = @(([System.Net.Dns]::GetHostAddresses($global:PCName)).IPAddressToString)
	if ($ips.length -eq 1){
		$global:IP = $ips[0]
		$IPAddressSourceLabel.Text ="(dns)"
		$true
	}elseif($ips.length -gt 1){
		$returnval = $false
		MakeClickList ([ref]$global:IP) $ips ([ref]$returnval)
		if(!$returnval){$false}else{$true}
	}
}
function pingIP{
	#$Online = Test-Connection -Computername $global:IP -BufferSize 16 -Count 1 -quiet
	#if($Online){ $IPAddressTextbox.Forecolor = "green" } else{ $IPAddressTextbox.Forecolor = "red" }
	$PingTimer.Enabled = $true
	$global:PingJob = Start-Job {
		param($IP)
		Test-Connection -Computername $IP -BufferSize 16 -Count 1 -quiet
	} -Arg $global:IP
}
 # fix later if WSMAN is broken that the ping check will fail.
 # Make WSMan Test and fix
function checkPing{
	if (($global:PingJob.State -eq "Completed")){
		$Pingable = @(Receive-Job -id $PingJob.id)
		if ($Pingable){
			$IPAddressTextbox.Forecolor = "Green"
		}else{
			$IPAddressTextbox.Forecolor = "Red"
		}
		$PingTimer.Enabled = $false
	}
}

### BUTTON CLICK FORMS ###
function extraUserFacts{
	$user = Get-ADUser $UserID -properties LockedOut, Enabled, AccountExpirationDate, Certificates, Department, Description, PasswordNeverExpires, BadPwdCount, LastBadPasswordAttempt, PasswordLastSet, WhenChanged, WhenCreated
	
	$UserFactsForm = createForm "ACtive Directory User Info" 300 350 "CenterScreen" "Fixed3D" $false $false $true
	#$UserFactsForm.Add_KeyDown({subWindowKeyListener($UserFactsForm)})
	$EnabledUserLabel                          	= New-Object system.Windows.Forms.Label
	$EnabledUserLabel.text                    	= ("Enabled: "  + $user.Enabled)
	$EnabledUserLabel.width                    	= 280
	$EnabledUserLabel.height                   	= 20
	$EnabledUserLabel.location                 	= New-Object System.Drawing.Point(10,10)
	$EnabledUserLabel.Font                     	= 'Microsoft Sans Serif,8.25'
	#$EnabledLbl = createItem "Label" 10 10 280 20 ("Enabled: "  + $user.Enabled) $UserFactsForm
	$UserAccountExpireDateLabel              	= New-Object system.Windows.Forms.Label
	$UserAccountExpireDateLabel.text         	= ("Account Expires: " + $user.AccountExpirationDate)
	$UserAccountExpireDateLabel.width       	= 280
	$UserAccountExpireDateLabel.height   		= 20
	$UserAccountExpireDateLabel.location     	= New-Object System.Drawing.Point(10,40)
	$UserAccountExpireDateLabel.Font   			= 'Microsoft Sans Serif,8.25'	
	#$UserAccountExpireDateLabel = createItem "Label" 10 40 280 20 ("Account Expires: " + $user.AccountExpirationDate) $UserFactsForm
	$PasswordNeverExpiresLabel        			= New-Object system.Windows.Forms.Label
	$PasswordNeverExpiresLabel.text        		= ("Password Never Expires: " + $($user.PasswordNeverExpires))
	$PasswordNeverExpiresLabel.width  			= 280
	$PasswordNeverExpiresLabel.height  			= 20
	$PasswordNeverExpiresLabel.location 		= New-Object System.Drawing.Point(10,70)
	$PasswordNeverExpiresLabel.Font       		= 'Microsoft Sans Serif,8.25'
	#$PasswordNeverExpiresLabel = createItem "Label" 10 70 280 20 ("Password Never Expires: " + $user.PasswordNeverExpires) $UserFactsForm
	$BadPasswordCountLabel            			= New-Object system.Windows.Forms.Label
	$BadPasswordCountLabel.text   				= ("Bad Password Count: " + $user.BadPwdCount)
	$BadPasswordCountLabel.width  				= 280
	$BadPasswordCountLabel.height  				= 20
	$BadPasswordCountLabel.location  			= New-Object System.Drawing.Point(10,100)
	$BadPasswordCountLabel.Font         		= 'Microsoft Sans Serif,8.25'	
	#$BadPasswordCountLabel = createItem "Label" 10 100 280 20 ("Bad Password Count: " + $user.BadPwdCount) $UserFactsForm
	$UserAccountCreationTimeLabel   			= New-Object system.Windows.Forms.Label
	$UserAccountCreationTimeLabel.text  		= ("Created Time: " + $user.whenCreated)
	$UserAccountCreationTimeLabel.width   		= 280
	$UserAccountCreationTimeLabel.height     	= 20
	$UserAccountCreationTimeLabel.location  	= New-Object System.Drawing.Point(10,130)
	$UserAccountCreationTimeLabel.Font       	= 'Microsoft Sans Serif,8.25'	
	#$UserAccountCreationTimeLabel = createItem "Label" 10 130 280 20 ("Created Time: " + $user.whenCreated) $UserFactsForm
	$UserAccountModifiedTimeLabel          		= New-Object system.Windows.Forms.Label
	$UserAccountModifiedTimeLabel.text     		= ("Modified Time: " + $user.whenChanged)
	$UserAccountModifiedTimeLabel.width    		= 280
	$UserAccountModifiedTimeLabel.height   		= 20
	$UserAccountModifiedTimeLabel.location    	= New-Object System.Drawing.Point(10,160)
	$UserAccountModifiedTimeLabel.Font         	= 'Microsoft Sans Serif,8.25'
	#$UserAccountModifiedTimeLabel = createItem "Label" 10 160 280 20 ("Modified Time: " + $user.whenChanged) $UserFactsForm
	$PasswordAgeLabel                         	= New-Object system.Windows.Forms.Label
	$PasswordAgeLabel.text                     	= (getPasswordAge)
	$PasswordAgeLabel.width                    	= 280
	$PasswordAgeLabel.height                   	= 20
	$PasswordAgeLabel.location                 	= New-Object System.Drawing.Point(10,190)
	$PasswordAgeLabel.Font                     	= 'Microsoft Sans Serif,8.25'	
	#$PasswordAgeLabel = createItem "Label" 10 190 280 20 (getPasswordAge) $UserFactsForm
	$CertificateButton                         	= New-Object system.Windows.Forms.Button
	$CertificateButton.text                    	= ""
	$CertificateButton.width                   	= 10
	$CertificateButton.height                  	= 20
	$CertificateButton.location                	= New-Object System.Drawing.Point(2,217)
	$CertificateButton.Font                    	= 'Microsoft Sans Serif,8.25'
	$CertificateButton.Visible = $False
	$global:CertStepper = 0
	$CertificateButton.Add_Click({ stepCert })
	#$CertificateButton = createItem "Button" 2 217 10 20 "" $UserFactsForm
	$CertificateLabel                        	= New-Object system.Windows.Forms.Label
	$CertificateLabel.text                    	= (getCertText).tostring()
	$CertificateLabel.width                    	= 270
	$CertificateLabel.height                   	= 20
	$CertificateLabel.location                 	= New-Object System.Drawing.Point(10,220)
	$CertificateLabel.Font                     	= 'Microsoft Sans Serif,8.25'
	#$certificateButton = createItem "Label" 10 220 270 20 (getCertText) $UserFactsForm
	$UserDepartmentLabel                      	= New-Object system.Windows.Forms.Label
	$UserDepartmentLabel.text               	= ("Department: " + $user.Department)
	$UserDepartmentLabel.width              	= 280
	$UserDepartmentLabel.height         		= 20
	$UserDepartmentLabel.location       		= New-Object System.Drawing.Point(10,250)
	$UserDepartmentLabel.Font              		= 'Microsoft Sans Serif,8.25'	
	#$UserDepartmentLabel = createItem "Label" 10 250 280 20 ("Department: " + $user.Department) $UserFactsForm
	$UserDescription				 			= New-Object system.Windows.Forms.Label
	$UserDescription.text                    	= ("Description: " + $user.Description)
	$UserDescription.width                    	= 280
	$UserDescription.height                   	= 20
	$UserDescription.location                 	= New-Object System.Drawing.Point(10,280)
	$UserDescription.Font                     	= 'Microsoft Sans Serif,8.25'	
	#$UserDescription = createItem "Label" 10 280 280 20 ("Description: " + $user.Description) $UserFactsForm

	$UserFactsForm.controls.AddRange(@($EnabledUserLabel,$UserAccountExpireDateLabel,$PWNeverExpireLb,
	$BadPasswordCountLabel,$UserAccountCreationTimeLabel,$UserAccountModifiedTimeLabel,$PasswordAgeLabel,$CertificateButton,$certificateButton,
	$UserDepartmentLabel,$UserDescription,$PasswordNeverExpiresLabel,$CertificateLabel))

	showForm $UserFactsForm
}
function extraPCFacts{
	$pc = Get-ADComputer $PCName -properties CanonicalName, Enabled, LastWrenchLogoPictureBoxnDate, OperatingSystem, OperatingSystemServicePack, WhenChanged, WhenCreated
	
	#Global Variables
	$global:MACStepper = 0
	$DetailsTimer = New-Object System.Windows.Forms.Timer
		$DetailsTimer.Interval = 1000
		$DetailsTimer.add_tick({ testDetailsData })
	$LoggedIn = ""
	$PCType = ""
	$MACs = ""
	$MemInfo = ""
	$global:flashline = ""
	$global:javaline = "" 
	$global:flashlinecount = 0
	$global:javalinecount = 0
	$global:softwaretimer = New-Object System.Windows.Forms.Timer
		$global:softwaretimer.Interval = 1000
	$MACsEdited = New-Object System.Collections.ArrayList
	
	# PC Facts Form
	$PCFactsForm = createForm "PC Facts" 300 550 "CenterScreen" "Fixed3D" $false $false $true
		$PCFactsForm.add_FormClosing({ $global:softwaretimer.enabled = $false; $detailstimer.enabled = $false })
		
		
	# AD info
	$ADUserEnabledlbl                          	= New-Object system.Windows.Forms.Label
	$ADUserEnabledlbl.text                    	= ("Enabled: "  + $pc.Enabled)
	$ADUserEnabledlbl.width                    	= 270
	$ADUserEnabledlbl.height                   	= 30
	$ADUserEnabledlbl.location                 	= New-Object System.Drawing.Point(10,10)
	$ADUserEnabledlbl.Font                     	= 'Microsoft Sans Serif,8.25'

	#$ADUserEnabledlbl = createItem "Label" 10 10 270 20 ("Enabled: "  + $pc.Enabled) $PCFactsForm
	$OULabel                         			= New-Object system.Windows.Forms.Label
	$OULabel.text                    			= ("OU: " + ($pc.CanonicalName.Substring(($pc.CanonicalName).IndexOf('/')))) 
	$OULabel.width                    			= 270
	$OULabel.height                   			= 20
	$OULabel.location                 			= New-Object System.Drawing.Point(10,40)
	$OULabel.Font                     			= 'Microsoft Sans Serif,8.25'
	#$OULabel = createItem "Label" 10 40 270 30 ("OU: " + ($pc.CanonicalName.Substring(($pc.CanonicalName).IndexOf('/')))) $PCFactsForm
    $LastWrenchLogoPictureBoxnLbl         		= New-Object system.Windows.Forms.Label
	$LastWrenchLogoPictureBoxnLbl.text     		= "Name: "
	$LastWrenchLogoPictureBoxnLbl.width    		= 270
	$LastWrenchLogoPictureBoxnLbl.height   		= 20
	$LastWrenchLogoPictureBoxnLbl.location 		= New-Object System.Drawing.Point(10,100)
	$LastWrenchLogoPictureBoxnLbl.Font     		= 'Microsoft Sans Serif,8.25'
	#$LastWrenchLogoPictureBoxnLbl = createItem "Label" 10 70 270 20 ("Last WrenchLogoPictureBoxn: " + $pc.LastWrenchLogoPictureBoxnDate) $PCFactsForm
	$OSLbl                          			= New-Object system.Windows.Forms.Label
	$OSLbl.text                     			= ("OS: " + $pc.OperatingSystem + " " + $pc.OperatingSystemServicePack)
	$OSLbl.width                    			= 270
	$OSLbl.height                   			= 20
	$OSLbl.location                 			= New-Object System.Drawing.Point(10,100)
	$OSLbl.Font                     			= 'Microsoft Sans Serif,8.25'
	#$OSLbl = createItem "Label" 10 100 270 20 ("OS: " + $pc.OperatingSystem + " " + $pc.OperatingSystemServicePack) $PCFactsForm
	$UserAccountCreationTimeLabel           	= New-Object system.Windows.Forms.Label
	$UserAccountCreationTimeLabel.text       	= ("Created Time: " + $pc.whenCreated)
	$UserAccountCreationTimeLabel.width         = 270
	$UserAccountCreationTimeLabel.height      	= 20
	$UserAccountCreationTimeLabel.location      = New-Object System.Drawing.Point(10,130)
	$UserAccountCreationTimeLabel.Font          = 'Microsoft Sans Serif,8.25'
	#$UserAccountCreationTimeLabel = createItem "Label" 10 130 270 20 ("Created Time: " + $pc.whenCreated) $PCFactsForm
	$UserAccountModifiedTimeLabel          		= New-Object system.Windows.Forms.Label
	$UserAccountModifiedTimeLabel.text     		= ("Modified Time: " + $pc.whenChanged)
	$UserAccountModifiedTimeLabel.width   		= 60
	$UserAccountModifiedTimeLabel.height   		= 20
	$UserAccountModifiedTimeLabel.location  	= New-Object System.Drawing.Point(10,28)
	$UserAccountModifiedTimeLabel.Font         	= 'Microsoft Sans Serif,8.25'
	#$UserAccountModifiedTimeLabel = createItem "Label" 10 160 270 20 ("Modified Time: " + $pc.whenChanged) $PCFactsForm

	# WMI info
	$UnderlineFont = New-Object System.Drawing.Font("Microsoft Sans Serif",8.5,[System.Drawing.FontStyle]::Underline)
	$NonUnderlineFont = New-Object System.Drawing.Font("Microsoft Sans Serif",8.5)
	$LoggedUserLbl = createItem "Label" 10 190 270 20 "Logged in user: " $PCFactsForm
		$LoggedUserLbl.Add_Click({ populateLoggedUser })
		$LoggedUserLbl.ForeColor = "darkblue"
		$LoggedUserLbl.Font = $UnderlineFont
	$PCTypeLbl = createItem "Label" 10 220 270 30 "PC Type: " $PCFactsForm
		$PCTypeLbl.Add_Click({ populatePCType })
		$PCTypeLbl.ForeColor = "darkblue"
		$PCTypeLbl.Font = $UnderlineFont
	$MACLbl = createItem "Label" 10 250 270 20 "MAC: " $PCFactsForm
		$MACLbl.Add_Click({ populateMAC })
		$MACLbl.ForeColor = "darkblue"
		$MACLbl.Font = $UnderlineFont
	$MACBtn = createItem "Button" 2 247 8 20  "" $PCFactsForm
		$MACBtn.Visible = $false
		$MACBtn.Add_Click({ stepMAC })
	$MemLbl = createItem "Label" 10 280 270 20 "Memory: " $PCFactsForm
		$MemLbl.Add_Click({ populateMem })
		$MemLbl.ForeColor = "darkblue"
		$MemLbl.Font = $UnderlineFont
	$DriveLbl = createItem "Label" 10 310 270 20 "Drive: " $PCFactsForm
		$DriveLbl.Add_Click({ populateDrive })
		$DriveLbl.ForeColor = "darkblue"
		$DriveLbl.Font = $UnderlineFont
	
	# PSinfo info
	$SoftVerButton                         		= New-Object system.Windows.Forms.Button
	$SoftVerButton.text                    		= "Uptime and Software Versions"
	$SoftVerButton.width                   		= 270
	$SoftVerButton.height                  		= 20
	$SoftVerButton.location                		= New-Object System.Drawing.Point(10,340)
	$SoftVerButton.Font                    		= 'Microsoft Sans Serif,8.25'
	$SoftVerButton.Add_Click({ getSoftwareVersions })
	#$SoftVerButton = createItem "Button" 10 340 270 20 "Uptime and Software Versions" $PCFactsForm
	$SoftVerLbl                          		= New-Object system.Windows.Forms.Label
	$SoftVerLbl.text                     		= "Loading ... (May take a while)"
	$SoftVerLbl.width                    		= 370
	$SoftVerLbl.height                   		= 20
	$SoftVerLbl.location                 		= New-Object System.Drawing.Point(10,340)
	$SoftVerLbl.Font                     		= 'Microsoft Sans Serif,8.25'
	$SoftVerLbl.ForeColor = "Blue"
	$SoftVerLbl.Visible = $false
	#$SoftVerLbl = createItem "Label" 10 340 270 20 "Loading ... (May take a while)" $PCFactsForm
	$UptimeLbl                          		= New-Object system.Windows.Forms.Label
	$UptimeLbl.text                     		= "Uptime:"
	$UptimeLbl.width                    		= 300
	$UptimeLbl.height                   		= 20
	$UptimeLbl.location                 		= New-Object System.Drawing.Point(10,340)
	$UptimeLbl.Font                     		= 'Microsoft Sans Serif,8.25'
	#$UptimeLbl = createItem "Label" 10 340 300 20 "Uptime:" $PCFactsForm
	$IELbl                          			= New-Object system.Windows.Forms.Label
	$IELbl.text                     			= "IE:"
	$IELbl.width                    			= 270
	$IELbl.height                   			= 20
	$IELbl.location                 			= New-Object System.Drawing.Point(10,370)
	$IELbl.Font                     			= 'Microsoft Sans Serif,8.25'
	#$IELbl = createItem "Label" 10 370 270 20 "IE:" $PCFactsForm
	$FlashLbl                         			= New-Object system.Windows.Forms.Label
	$FlashLbl.text                     			= "Flash:"
	$FlashLbl.width                    			= 270
	$FlashLbl.height                   			= 20
	$FlashLbl.location                 			= New-Object System.Drawing.Point(10,400)
	$FlashLbl.Font                     			= 'Microsoft Sans Serif,8.25'	
	#$FlashLbl = createItem "Label" 10 400 270 20 "Flash:" $PCFactsForm
	$FlashButton                         		= New-Object system.Windows.Forms.Button
	$FlashButton.text                    		= ""
	$FlashButton.width                   		= 8
	$FlashButton.height                  		= 20
	$FlashButton.location                		= New-Object System.Drawing.Point(2,397)
	$FlashButton.Font                    		= 'Microsoft Sans Serif,8.25'
	$FlashButton.Visible = $false
	$FlashButton.Add_Click({ stepFlash })	
	#$FlashButton = createItem "Button" 2 397 8 20 "" $PCFactsForm
	$JavaLbl.text                     			= "Java:"
	$JavaLbl.width                    			= 270
	$JavaLbl.height                   			= 20
	$JavaLbl.location                 			= New-Object System.Drawing.Point(2,247)
	$JavaLbl.Font                     			= 'Microsoft Sans Serif,8.25'
	#$JavaLbl = createItem "Label" 10 430 270 20 "Java:" $PCFactsForm
	$JavaButton.text                    		= ""
	$JavaButton.width                   		= 8
	$JavaButton.height                  		= 20
	$JavaButton.location                		= New-Object System.Drawing.Point(2,427)
	$JavaButton.Font                    		= 'Microsoft Sans Serif,8.25'
	$JavaButton.Visible = $false
	$JavaButton.Add_Click({ stepJava })	
	#$JavaButton = createItem "Button" 2 427 8 20 "" $PCFactsForm
	
	#Endpoint Protection
	$EndPointButton                         	= New-Object system.Windows.Forms.Button
	$EndPointButton.text                    	= "View Endpoint Details"
	$EndPointButton.width                   	= 270
	$EndPointButton.height                  	= 20
	$EndPointButton.location                	= New-Object System.Drawing.Point(10,460)
	$EndPointButton.Font                   		= 'Microsoft Sans Serif,8.25'	
	$EndPointButton = createItem "Button" 10 460 270 20 "View Endpoint Details" $PCFactsForm
	$EndPointButton.Add_Click({getEndpointInfo $EndPointButton $EndpointLabel})
	
	$EndPointLabel                          	= New-Object system.Windows.Forms.Label
	$EndPointLabel.text                     	= "Signatures Last Updated: "
	$EndPointLabel.width                    	= 270
	$EndPointLabel.height                   	= 20
	$EndPointLabel.location                 	= New-Object System.Drawing.Point(10,460)
	$EndPointLabel.Font                     	= 'Microsoft Sans Serif,8.25'
	$EndPointLabel.Visible = $False	
	#$EndPointLabel = createItem "Label" 10 460 270 20 "Signatures Last Updated: " $PCFactsForm
	#View Smart Data
	$SmartLbl                          			= New-Object system.Windows.Forms.Label
	$SmartLbl.text                     			= "View Disk Health"
	$SmartLbl.width                    			= 270
	$SmartLbl.height                  			= 20
	$SmartLbl.location                 			= New-Object System.Drawing.Point(100,490)
	$SmartLbl.Font                     			= 'Microsoft Sans Serif,8.25'
	$SmartLbl.ForeColor = "Blue"
	$SmartLbl.Add_Click({ getSmartData })	
	#$SmartLbl = createItem "Label" 100 490 270 20 "View Disk Health" $PCFactsForm
	$PCFactsForm.controls.AddRange(@($SoftVerButton,$SoftVerLbl,$UptimeLbl,$IELbl,$FlashLb,$FlashButton,$JavaLbl,$JavaButton,$EndPointButton,$EndPointLabel,$SmartLbl))
	showForm $PCFactsForm
}
function viewGroups($ADObjectType){
		$GroupForm = createForm "Groups" 250 300 "CenterScreen" "Fixed3D" $false $false $true

		$GroupList                        		= New-Object system.Windows.Forms.ListBox
		$GroupList.text                   		= "" # Starts Empty. Is this needed.
		$GroupList.width                 		= 210
		$GroupList.height                 		= 180
		$GroupList.location               		= New-Object System.Drawing.Point(10,10)
		$GroupList.HorizontalScrollbar 			= $true
		$GroupList.DataSource = @((getGroups($ADObjectType)).SamAccountName | Sort-Object )
		#$GroupList = createItem "ListBox" 10 10 210 180 "" $GroupForm

		$AddGroupButton                         = New-Object system.Windows.Forms.Button
		$AddGroupButton.text                    = "Add to Group"
		$AddGroupButton.width                   = 210
		$AddGroupButton.height                  = 30
		$AddGroupButton.location                = New-Object System.Drawing.Point(10,190)
		$AddGroupButton.Font                    = 'Microsoft Sans Serif,8.25'
		$AddGroupButton.Add_Click({ addGroup $ADObjectType})	
		#$AddGroupButton = createItem "Button" 10 190 210 30 "Add to Group" $GroupForm
		
		$RemoveGroupButton                  	= New-Object system.Windows.Forms.Button
		$RemoveGroupButton.text            		= "Remove from Group"
		$RemoveGroupButton.width           		= 210
		$RemoveGroupButton.height          		= 30
		$RemoveGroupButton.location      		= New-Object System.Drawing.Point(10,225)
		$RemoveGroupButton.Font             	= 'Microsoft Sans Serif,8.25'
		$RemoveGroupButton.Add_Click({ removeGroup $ADObjectType})
		#$RemoveGroupButton = createItem "Button" 10 225 210 30 "Remove from Group" $GroupForm
		$GroupForm.controls.AddRange(@($GroupList,$AddGroupButton,$RemoveGroupButton))

		showForm($GroupForm)
}
function newUserPassword{
	$newPWForm = createForm "New Password" 300 160 "CenterScreen" "Fixed3D" $false $false $false
	$NewPasswordLabel                         			= New-Object system.Windows.Forms.Label
	$NewPasswordLabel.text                     	= "Enter New Password: "
	$NewPasswordLabel.width                    	= 120
	$NewPasswordLabel.height                   	= 20
	$NewPasswordLabel.location                 	= New-Object System.Drawing.Point(10,12)
	$NewPasswordLabel.Font                     	= 'Microsoft Sans Serif,8.25'
	# $NewPasswordLabell = createItem "Label" 1	0 12 120 20 "Enter New Password: " $newPWForm
	$NewPasswordTextbox                        	= New-Object system.Windows.Forms.TextBox
	$NewPasswordTextbox.width                  	= 120
	$NewPasswordTextbox.height                 	= 20
	$NewPasswordTextbox.location               	= New-Object System.Drawing.Point(140,10)
	$NewPasswordTextbox.Font                   	= 'Microsoft Sans Serif,8.25'
	$NewPasswordTextbox.PasswordChar 			= "*"
	#$NewPasswordTextbox = createItem "TextBox" 140 10 120 20 "" $newPWForm
	$ConfirmPasswordLabel                		= New-Object system.Windows.Forms.Label
	$ConfirmPasswordLabel.text          		= "Name: "
	$ConfirmPasswordLabel.width          		= 120
	$ConfirmPasswordLabel.height          		= 20
	$ConfirmPasswordLabel.location       		= New-Object System.Drawing.Point(10,42)
	$ConfirmPasswordLabel.Font          		= 'Microsoft Sans Serif,8.25'	
	#$ConfirmPasswordLabel = createItem "Label" 10 42 120 20 "Confirm Password: " $newPWForm
	$ConfirmPaswordTextbox            			= New-Object system.Windows.Forms.TextBox
	$ConfirmPaswordTextbox.width          		= 120
	$ConfirmPaswordTextbox.height       		= 20
	$ConfirmPaswordTextbox.location           	= New-Object System.Drawing.Point(140,40)
	$ConfirmPaswordTextbox.Font       			= 'Microsoft Sans Serif,8.25'
	$ConfirmPaswordTextbox.PasswordChar 		= "*"
	#$ConfirmPaswordTextbox = createItem "TextBox" 140 40 120 20 "" $newPWForm
	$ChangePasswordCheckbox                   	= New-Object system.Windows.Forms.CheckBox
	$ChangePasswordCheckbox.text              	= "checkBox"
	$ChangePasswordCheckbox.AutoSize          	= $false
	$ChangePasswordCheckbox.width              	= 220
	$ChangePasswordCheckbox.height             	= 20
	$ChangePasswordCheckbox.location           	= New-Object System.Drawing.Point(60,67)
	$ChangePasswordCheckbox.Font                = 'Microsoft Sans Serif,10'	
	#$ChangePasswordCheckbox = createItem "CheckBox" 60 67 220 20 "Change password at next login" $newPWForm
	$NewPasswordOkButton                        = New-Object system.Windows.Forms.Button
	$NewPasswordOkButton.text                   = "OK"
	$NewPasswordOkButton.width                  = 122
	$NewPasswordOkButton.height                 = 20
	$NewPasswordOkButton.location               = New-Object System.Drawing.Point(10,90)
	$NewPasswordOkButton.Font                   = 'Microsoft Sans Serif,8.25'
	#$NewPasswordOkButton = createItem "Button" 10 90 122 20 "OK" $newPWForm
	$NewPasswordCancelButton                    = New-Object system.Windows.Forms.Button
	$NewPasswordCancelButton.text               = "Cancel"
	$NewPasswordCancelButton.width              = 122
	$NewPasswordCancelButton.height             = 20
	$NewPasswordCancelButton.location           = New-Object System.Drawing.Point(147,90)
	$NewPasswordCancelButton.Font               = 'Microsoft Sans Serif,8.25'
	#$NewPasswordCancelButton = createItem "Button" 147 90 122 20 "Cancel" $newPWForm

	$newPWForm.controls.AddRange(@($NewPasswordLabel,$NewPasswordTextbox,$ConfirmPasswordLabel,$ConfirmPaswordTextbox,$ChangePasswordCheckbox,$NewPasswordOkButton,$NewPasswordCancelButton))

	$NewPasswordOkButton.Add_Click({
		if ($NewPasswordTextbox.Text -eq $ConfirmPaswordTextbox.text){
			try{
				$pw = ConvertTo-SecureString $NewPasswordTextbox.Text -AsPlainText -Force
				Set-ADAccountPassword $global:UserID -Reset -NewPassword $pw	
				if ($ChangePasswordCheckbox.Checked -eq $true){
					Set-ADUser $global:UserID -ChangePasswordAtWrenchLogoPictureBoxn:$true
				}
				$newPWForm.Close()
			}catch{
				$msg::Show($error[0])
			}
		}else{
			$msg::Show("Passwords do not match")
		}
	})
	
	$NewPasswordCancelButton.Add_Click({
		$NewPWForm.Close()
	})
	
	$newPWForm.Add_Shown({$newPWForm.Activate()})	
	[void] $newPWForm.ShowDialog()
	$newPWForm.BringToFront()	
}
function viewCFolder{
	$cpath = '\\' + $global:IP + '\c$'
	Start-process explorer.exe $cpath
}
function telnetPC{
	Start-Process telnet $global:IP
}
function expandForm{
	if ($mainForm.Width -eq 300){
		$mainForm.Size = New-Object System.Drawing.Size(445,635)
		$ExpandButton.Text = "<"
	}else{
		$mainForm.Size = New-Object System.Drawing.Size(300,635)
		$ExpandButton.Text = ">"
	}
}

### USER FACT FUNCTIONS ###
function getPasswordAge{
	$CurrentTime = Get-Date
	$PWChangedTime = $user.PasswordLastSet
	$PassAgeText = ""
	try{
		$TimeDifference = New-TimeSpan -Start $PWChangedTime -End $CurrentTime
		$PassAgeText = "Password Age: " + $TimeDifference.Days + " Days and " + $TimeDifference.Hours + " Hours"
	}catch{
		$PassAgeText = "Password Age: Unavailable"
	}
	$PassAgeText
}
function getCertText{
	if (($user.certificates).count -lt 1){
		"Certificate: None"
	}elseif(($user.certificates).count -eq 1){
		listCert(0)
	}else{
		$CertificateButton.Visible = $true
		listCert(0)
	}
}
function listCert($certstep){
	$StartDate = $user.Certificates[$certstep].GetEffectiveDateString()
	$StartDate = $StartDate.Substring(0,$StartDate.IndexOf(" "))
	$EndDate = $user.Certificates[$certstep].GetExpirationDateString()
	$EndDate = $EndDate.Substring(0,$EndDate.IndexOf(" "))
	"Certificate: " + $StartDate + " - " + $EndDate
}
function stepCert{
	$global:CertStepper++
	$index = ($CertStepper%($user.certificates).count)
	$certificateButton.Text = listCert($index)
}

### PC FACT FUNCTIONS ###
function getSmartData{
	$CommandString = "/c start cmd /k " + $PSExecLocation + " -accepteula \\$global:PCName " + $SmartCtlLocation + " -a sda"
	Start-Process "cmd.exe" $CommandString
}
function testDetailsData{
	if($global:LoggedInJob.State -eq "Completed"){
		$LoggedIn = @(Receive-Job -id $LoggedInJob.id)
		displayDetails $LoggedIn "LoggedIn"
		$global:LoggedInJob = ""
	}
	if($global:PCTypeJob.State -eq "Completed"){
		$PCType = @(Receive-Job -id $PCTypeJob.id)
		displayDetails $PCType "PCType"
		$global:PCTypeJob = ""
	}
	if($global:MACJob.State -eq "Completed"){
		$MACs = @(Receive-Job -id $MACJob.id)
		displayDetails $MACs "MAC"
		$global:MACJob = ""
	}
	if($global:MemJob.State -eq "Completed"){
		$MemInfo = Receive-Job -id $MemJob.id
		displayDetails $MemInfo "Memory"
		$global:MemJob = ""
	}
	if($global:DriveJob.State -eq "Completed"){
		$DriveInfo = Receive-Job -id $DriveJob.id
		displayDetails $DriveInfo "Drive"
		$global:DriveJob = ""
	}
}
function displayDetails($jobresult, $name){
	if (($global:LoggedInJob.State -ne "Running") -and ($global:PCTypeJob.State -ne "Running") -and ($global:MACJob.State -ne "Running") -and ($global:MemJob.State -ne "Running") -and ($global:DriveJob.State -ne "Running")){
		$DetailsTimer.enabled = $false
	}
	switch($name){
		"LoggedIn" {
			$LoggedUserLbl.Text = "Logged in user: " + $jobresult   
		}
		"PCType"{
			$PCTypeLbl.Text = "PC Type: " + $jobresult
		}
		"MAC"{
			for ($i = 1; $i -lt $jobresult.Count; $i++){
				$jobresult[$i] = $jobresult[$i].Substring(0,17)
				if($jobresult[$i] -ne "" -and $jobresult[$i] -ne "N/A              "){
					$MACsEdited.Add($jobresult[$i])	
				}
			}
			$MACLbl.Text = "MAC: " + $MACsEdited[0]
			if ($MACsEdited.count -gt 1){
				$MACBtn.Visible = $True
			}
		}
		"Memory"{
			$MemLbl.Text = "Memory: " + $jobresult + " MHz"
		}
		"Drive"{
			$DriveLbl.Text = "Drive: " + $jobresult
		}
	}
	
}
function testSoftData{
	if ($global:SoftJob.State -eq "Completed"){
		$global:Softwares = @(Receive-Job -id $global:SoftJob.ID)
		displaySoftware
	}
}
function displaySoftware{
	try{
		$global:softwaretimer.enabled = $false
		$IEVersion = getIEVersion
		$UptimeLine = ([String]($global:softwares | select-string "Uptime")).Substring(26)
		$global:FlashLine = @($global:softwares | select-string "Adobe Flash Player")
		for($i = 0; $i -lt $flashline.count; $i++){
			$flashline[$i] = [string]$flashline[$i]
			$flashline[$i] = $flashline[$i].Substring($flashline[$i].IndexOf("ActiveX"))
		}
		$global:JavaLine = @($global:softwares | select-string "Java")
		$IELbl.Text = "IE:            " + $IEVersion
		$FlashLbl.Text = "Flash:      " + $flashline[0]
		$JavaLbl.Text = "Java:       " + $javaline[0]
		$UptimeLbl.Text = "Uptime:    " + $UptimeLine
		if ($FlashLine.count -gt 1){
			$FlashButtonIndex = 0
			$FlashButton.Visible = $true

		}
		if ($JavaLine.count -gt 1){
			$JavaButtonIndex = 0
			$JavaButton.Visible = $true
		}
		$IELbl.Visible = $true
		$FlashLbl.Visible = $true
		$Javalbl.Visible = $true
		$SoftVerLbl.Visible = $false
		
	}catch{
		$msg::Show($error[0])
	}
}
function getIEVersion{
		$IEReg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $global:pcname)
		$IERegKey= $IEReg.OpenSubKey("SOFTWARE\\Microsoft\\Internet Explorer\\")
		$IEVersion = $IERegKey.GetValue("svcVersion")
		if ($IEVersion.count -gt 0){
			$IEVersion
		}else{
			$IEVersion = $IERegKey.GetValue("Version")
			$IEVersion
		}
}
function getSoftwareVersions{
	$global:softwaretimer.add_tick({ testSoftData })
	$SoftVerButton.Visible = $False
	$SoftVerLbl.Visible = $true
	$Online = Test-Connection -Computername $global:PCName -BufferSize 16 -Count 1 -quiet
	if($Online){
		$cmdblk = "& $PSInfoLocation `$args"
		$commandblock = [scriptblock]::Create($cmdblk)
		$global:SoftJob = Start-Job -scriptblock $commandblock -argumentlist "\\$global:pcname", "-s"
		
		$global:softwaretimer.start()
		testSoftData
	}else{
		$SoftVerLbl.Visible = $false
		$msg::Show("Not Online")
	}
}
function stepMAC{
	$global:MACStepper = $global:MACStepper + 1
	$index = ($MACStepper%($MACsEdited.count))
	$MACLbl.Text = "MAC: " + $MACsEdited[$index]
}
function stepJava{
	$global:JavaButtonIndex++
	$JavaLbl.Text = "Java:      " + $JavaLine[$JavaButtonIndex % ($Javaline.count)]
}
function stepFlash{
	$global:FlashButtonIndex++
	$FlashLbl.Text = "Flash:      " + $flashline[$FlashButtonIndex % ($Flashline.count)]
}
function getEndpointInfo($Button, $Label){
	$Online = Test-Connection -Computername $global:IP -BufferSize 16 -Count 1 -quiet
	$PCOS = (Get-ADComputer $global:pcname -properties OperatingSystem).OperatingSystem
	if($Online){
		$Registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $global:PCName)
		if ($PCOS -Like "Windows 7*"){
			$RegistryKey= $Registry.OpenSubKey("SOFTWARE\\Microsoft\\Microsoft Antimalware\\Signature Updates")
		}elseif ($PCOS -Like "Windows 10*"){
			$RegistryKey= $Registry.OpenSubKey("SOFTWARE\\Microsoft\\Windows Defender\\Signature Updates")
		
		}
		$Data = $RegistryKey.GetValue("SignaturesLastUpdated")
		$Time = [DateTime]::FromFileTime( (((((($Data[7]*256 + $Data[6])*256 + $Data[5])*256 + $Data[4])*256 + $Data[3])*256 + $Data[2])*256 + $Data[1])*256 + $Data[0])
		$Label.Text = "Signatures Last Updated: " + $Time
		$Button.visible = $false
		$Label.visible = $true
	}else{
		$msg::Show("PC if offline")
	}

}
function populateLoggedUser{
	$online = Test-Connection -Computername $global:PCName -BufferSize 16 -Count 1 -quiet
	if($online){
		$global:LoggedInJob = Start-Job {
			param($PCName)
			@(Get-WmiObject -Class Win32_ComputerSystem -ComputerName $PCName -erroraction silentlycontinue)[0].UserName
		} -Arg $global:PCName
	}
	$LoggedUserLbl.Font = $NonUnderlineFont
	$LoggedUserLbl.ForeColor = "black"
	$LoggedUserLbl.Text = "Logged in user: Please wait..."
	$DetailsTimer.start()
}
function populatePCType{
	$online = Test-Connection -Computername $global:PCName -BufferSize 16 -Count 1 -quiet
	if($online){
		$global:PCTypeJob = Start-Job {
			param($PCName)
			(Get-WmiObject -class win32_ComputerSystemProduct -computername $PCName).version 
			" " 
			(Get-WmiObject -Class Win32_ComputerSystem -ComputerName $PCName).Model
			" "
			(Get-WmiObject -Class Win32_Bios -ComputerName $PCName | Select-Object @{Name='mfgDate' ; expression ={$_.ConvertToDateTime($_.ReleaseDate).ToShortDateString()}}).mfgDate
		} -Arg $global:PCName
	}
	$PCTypeLbl.Font = $NonUnderlineFont
	$PCTypeLbl.ForeColor = "black"
	$PCTypeLbl.Text = "PC Type: Please wait..."
	$DetailsTimer.start()
}
function populateMAC{
	$online = Test-Connection -Computername $global:PCName -BufferSize 16 -Count 1 -quiet
	if($online){
		$global:MACJob = Start-Job {
			param($PCName)
			getmac /s $PCName /nh
		} -Arg $global:PCName
	}
	$MACLbl.Font = $NonUnderlineFont
	$MACLbl.ForeColor = "black"
	$MACLbl.Text = "Mac: Please wait..."
	$DetailsTimer.start()
}
function populateMem{
	$online = Test-Connection -Computername $global:PCName -BufferSize 16 -Count 1 -quiet
	if($online){
		$global:MEMJob = Start-Job {
			param($PCName)
			((Get-WmiObject -Class Win32_PhysicalMemory -ComputerName $PCName) | Measure-Object -Property Capacity -Sum).Sum / 1GB
			" GB "
			(((Get-WmiObject -Class WIn32_PhysicalMemory -ComputerName $PCName).speed) | Measure-Object -min ).minimum
		} -Arg $global:PCName
	}
	$MemLbl.Font = $NonUnderlineFont
	$MemLbl.ForeColor = "black"
	$MemLbl.Text = "Memory: Please wait..."
	$DetailsTimer.start()
}
function populateDrive{
	$online = Test-Connection -Computername $global:PCName -BufferSize 16 -Count 1 -quiet
	if($online){
		$global:DriveJob = Start-Job {
			param($PCName)
			$disk = get-WmiObject win32_logicaldisk -ComputerName $PCName
			if($disk.DeviceID.count -gt 1){$disk = $disk | Where-Object DeviceID -like "C*"}
			$util = ($disk.Size - $disk.freespace) / $disk.size
			$utilstr = ("{0:N2}" -f ($util * 100)) + "%"
			$outstr = ("{0:N2}" -f (($disk.Size - $disk.freespace)/1000000000)) + "GB / " + ("{0:N2}" -f ($disk.size/1000000000)) + "GB (" + $utilstr + " Used)"
			$outstr
		} -Arg $global:PCName
	}
	$DriveLbl.Font = $NonUnderlineFont
	$DriveLbl.ForeColor = "black"
	$DriveLbl.Text = "Drive: Please wait..."
	$DetailsTimer.start()
}

### GROUP FUNCTIONS ###
function getObject($ADObjectType){
	if ($ADObjectType -eq "user"){
		$ADObject = Get-ADUser "$global:UserID" -Properties memberOf 
	}elseif ($ADObjectType -eq "computer"){
		$ADObject = Get-ADComputer "$global:PCName" -Properties memberOf 
	}
	$ADObject
}
function getGroups($ADObjectType){
	(getObject($ADObjectType)).memberOf | ForEach-Object {Get-ADGroup $_ }
}
function addObjectToGroup(){
	try{
		$ADObject = getObject $ADObjectType
		$Group = $AddGroupList.SelectedItem
		Add-ADGroupMember -Identity $Group -Members $ADObject
		updateGroups
		$addGroupForm.close()
	}catch{
		$msg::Show($error[0])
	}
}
function searchGroups(){
	$search = "*" + $addGroupSearchBox.Text + "*"
	$Groups = @((Get-ADGroup -f {name -like $search}).name | Sort-Object)
	$addGroupList.DataSource = $Groups
	$addGroupList.Focus()
}
function removeFromGroup{
	try{
		$Object = getObject($ADObjectType)
		$Group = $RemoveGroupList.SelectedItem
		Remove-ADGroupMember -Identity $group -Members $Object -Confirm:$false 
		updateGroups
		$removeGroupForm.close()
	}catch{
		$msg::Show($error[0])
	}
}
function groupKeyboard(){
	if (($_.KeyCode -eq "Enter") -and ($addGroupList.Focused -eq $True)){
		addObjectToGroup
	}
	if (($_.KeyCode -eq "Enter") -and ($addGroupSearchBox.Focused -eq $True)){
		searchGroups
	}
}
function addGroup($ADObjectType){
	$addGroupForm = createForm "Add to Group" 250 300 "CenterScreen" "Fixed3D" $false $false $true
		$addGroupForm.KeyPreview = $True
		$addGroupForm.Add_KeyDown({ groupKeyboard })
	$addGroupSearchBox = createItem "TextBox" 10 10 150 20 "" $addGroupForm
	$addGroupSearchButton = createItem "Button" 170 10 50 20 "Search" $addGroupForm
		$addGroupSearchButton.Add_Click({ searchGroups })
	$addGroupList = createItem "ListBox" 10 40 210 180 "" $addGroupForm
		$addGroupList.HorizontalScrollbar = $true
		$addGroupList.add_MouseDoubleClick({ addObjectToGroup })
	$addToGroupButton = createItem "Button" 10 220 210 30 "Add" $addGroupForm
		$addToGroupButton.Add_Click({addObjectToGroup})
		
	showForm $addGroupForm
}
function removeGroup($ADObjectType){
	$removeGroupForm = createForm "Remove from group" 250 300 "CenterScreen" "Fixed3D" $false $false $true
		$removeGroupForm.KeyPreview = $True
		$removeGroupForm.Add_KeyDown({ if ($_.KeyCode -eq "Enter") {removeFromGroup} })
	$RemoveGroupList = createItem "ListBox" 10 10 210 200 "" $removeGroupForm
		$RemoveGroupList.HorizontalScrollbar = $true
		$RemoveGroupList.Focus()
		$RemoveGroupList.add_MouseDoubleClick({ removeFromGroup	})
		$RemoveGroupList.DataSource =  @((getGroups($ADObjectType)).SamAccountName | Sort-Object)
	$RemoveFromGroupButton = createItem "Button" 10 220 210 30 "Remove" $removeGroupForm
		$RemoveFromGroupButton.Add_Click({removeFromGroup})

	showForm $removeGroupForm
}
function updateGroups{
	$GroupList.DataSource =  @((getGroups($ADObjectType)).SamAccountName)
	$GroupList.Update()
}

showForm $mainForm
}
