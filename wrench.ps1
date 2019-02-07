
###############################################################################
# Program Name: Wrench
# Author: Matt Tuchfarber
# Contributors: James Atkinson, Kevin Cook
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
	$NameLbl = createItem "Label" 10 28 60 20 "Name: " $mainForm
	$NameBox = createItem "TextBox" 70 25 130 20 "" $mainForm
	$NameButton = createItem "Button" 210 25 60 20 "Search" $mainForm
		$NameButton.TabStop = $False
		$NameButton.Add_Click({	searchByName })
	#UserID Info
	$UserIDLbl = createItem "Label" 10 58 60 20 "User ID: " $mainForm
	$UserIDBox = createItem "TextBox" 70 55 130 20 "" $mainForm
		#$UserIDBox.MaxLength = 15
	$UserIDButton = createItem "Button" 210 55 60 20 "Search" $mainForm
		$UserIDButton.TabStop = $False
		$UserIDButton.Add_Click({ searchByUserID })
	#PC Name Info
	$PCLbl = createItem "Label" 10 88 60 20 "PC Name: " $mainForm
	$PCBox = createItem "TextBox" 70 85 130 20 "" $mainForm
		$PCBox.MaxLength = 15
	$PCButton = createItem "Button" 210 85 60 20 "Search" $mainForm
		$PCButton.TabStop = $False
		$PCButton.Add_Click({ searchByPCName })
	#IP Info
	$IPLbl = createItem "Label" 10 118 18 20 "IP: " $mainForm
	$IPSourceLbl = createItem "Label" 28 118 40 20 "" $mainForm
	$IPBox = createItem "Textbox" 70 115 130 20 "" $mainForm
		$IPBox.MaxLength = 15
	$IPButton = createItem "Button" 210 115 60 20 "Search" $mainForm
		$IPButton.TabStop = $False
		$IPButton.Add_Click({ searchByIP })
	#Phone Info
	$PhoneLbl = createItem "Label" 10 148 60 20 "Phone:" $mainForm
	$PhoneBox = createItem "Textbox" 70 145 200 20 "" $mainForm
		$PhoneBox.ReadOnly = $True
		$PhoneBox.TabStop = $False
	#Lockout Info
	$LockoutLbl = createItem "Label" 10 178 60 20 "Lockout:" $mainForm
	$LockoutBox = createItem "Textbox" 70 175 130 20 "" $mainForm
		$LockoutBox.ReadOnly = $True
		$LockoutBox.TabStop = $False
	$LockoutButton = createItem "Button" 210 175 60 20 "Unlock" $mainForm
		$LockoutButton.TabStop = $False
		$LockoutButton.Visible = $False
		$LockoutButton.Add_Click({ unlockAccount })
	#H Drive Info
	$HDriveLbl = createItem "Label" 10 208 60 20 "H Drive:" $mainForm
	$HDriveBox = createItem "Textbox" 70 208 200 20 "" $mainForm
		$HDriveBox.ReadOnly = $True
		$HDriveBox.TabStop = $False
	#OU Info
	$OULbl = createItem "Label" 10 238 60 20 "User OU:" $mainForm
	$OUBox = createItem "Textbox" 70 235 200 20 "" $mainForm
		$OUBox.ReadOnly = $True
		$OUBox.TabStop = $False
	#Buttons
	$RVButton = createItem "Button" 10 265 259 20 "Connect Via SCCM Remote Control" $mainForm
		$RVButton.Add_Click({ runRemoteViewer })		
	$UserFactsButton = createItem "Button" 10 295 122 20 "User Details" $mainForm
		$UserFactsButton.Add_Click({ runUserFacts })
	$UserGroupButton = createItem "Button" 10 325 122 20 "User Groups" $mainForm
		$UserGroupButton.Add_Click({ runUserGroups })
	$ChangePWButton = createItem "Button" 10 355 122 20 "Change Password" $mainForm
		$ChangePWButton.Add_Click({newUserPassword})
	$PCFactsButton = createItem "Button" 147 295 122 20 "PC Details" $mainForm
		$PCFactsButton.Add_Click({ runPCFacts })
	$PCGroupButton = createItem "Button" 147 325 122 20 "PC Groups" $mainForm
		$PCGroupButton.Add_Click({ runPCGroups })
	$PCManageButton = createItem "Button" 147 355 122 20 "Manage PC" $mainForm
		$PCManageButton.Add_Click({ runManagePC })
	$ViewCButton = createItem "Button" 147 385 122 20 "View C:" $mainForm
		$ViewCButton.Add_Click({ runViewC })
	$RDPButton = createItem "Button" 147 415 122 20 "RDP" $mainForm
		$RDPButton.Add_Click({ runRDP })
	$PSSessionBtn = createItem "Button" 147 445 122 20 "PS Remote" $mainForm
		$PSSessionBtn.Add_Click({connectPSSession})
	$OnTopCheck = createItem "Checkbox" 30 410 100 35 "Keep Wrench on Top" $mainForm
		$OnTopCheck.Add_Click({ runOnTop })
	$NewPSLbl = createItem "Label" 80 578 150 15 "New Powershell Window" $mainForm
		$NewPSLbl.ForeColor = "Blue"
		$NewPSLbl.Add_Click({ Start-Process "powershell.exe" })
	$PingTimer = New-Object System.Windows.Forms.Timer
		$PingTimer.Interval = 1000
		$PingTimer.add_tick({ checkPing })

	#Logo Image
	$logo = new-object Windows.Forms.PictureBox
	$logo.location = New-Object System.Drawing.Size(10,477)
	$logo.size = New-Object System.Drawing.Size(260,100)
	$logo.BorderStyle = "FixedSingle"
	$logo.Image = [System.Drawing.Image]::Fromfile((get-item $LogoLocation));
	$mainForm.controls.add($logo)
	
	#Expand Button
	$ExpandButton = createItem "Button" 264 577 15 15 ">" $mainForm
		$ExpandButton.Add_Click({ expandForm })
		$ExpandButton.Visible = $true

	####### EXPANDED FORM GUI #######
	
	#Draw Seperator
	$pen = new-object Drawing.SolidBrush black
	$formGraphics = $mainForm.createGraphics()
	$mainForm.add_paint({$formGraphics.DrawLine($pen, 279, 1, 279, 620)})
		
	#Buttons
	$GPBtn = createItem "Button" 290 25 122 20 "Check Group Policy" $mainForm
		$GPBtn.Add_Click({ pullGroupPolicy })
	$GPTimer = New-Object System.Windows.Forms.Timer
		$GPTimer.Interval = 2000
		$GPTimer.add_tick({ checkGP })
	$TelnetPCButton = createItem "Button" 290 55 122 20 "Telnet" $mainForm
		$TelnetPCButton.Add_Click({ runTelnet })
	$RenameButton = createItem "Button" 290 85 122 20 "Rename PC" $mainForm
		$RenameButton.Add_Click({ runRename })
	$SCCMClientCenterButton = createItem "Button" 290 115 122 20 "Client Center" $mainForm
		$SCCMClientCenterButton.Add_Click({ openClientCenter })
	$MSRemoteAssistance = createItem "Button" 290 145 122 20 "MS Remote Assist" $mainForm
		$MSRemoteAssistance.Add_Click({ openMSRA })
	
	
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
		$global:Name = ($NameBox.text).Trim()
		testName		
		if ($global:validName -eq $True){
			if(!(pickName)){return}
			clearBoxes
			$NameBox.text = $global:Name
			$LockoutButton.Visible=$false
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
		$global:UserID = ($UserIDBox.Text).Trim()
		testID
		if($global:validID -eq $True){
			if(!(pickID)){return}
			clearBoxes
			$UserIDBox.Text = $global:UserID
			$LockoutButton.Visible=$false
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
	function searchByPCName{
		clearVariables
		$global:PCName = ($PCBox.Text).Trim()
		testPCName
		if($global:validPCName -eq $True){
			if(!(pickPCName)){return}
			clearBoxes
			$PCBox.Text = $global:PCName
			$LockoutButton.Visible=$false
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
		$global:IP = ($IPBox.Text).Trim()
		if(testIP){
			clearBoxes
			$IPBox.Text = $global:IP
			if($global:IPWithARec){
				$LockoutButton.Visible=$false
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
				$LockoutButton.Visible = $False
				$LockoutBox.Text = "Not Locked"
			}else{
				$LockoutBox.Text = "Locked"
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
		if ($OnTopCheck.Checked -eq $true){
			$mainForm.TopMost = $true
			$mainForm.Update()
		}else{
			$mainForm.TopMost = $false
			$mainForm.Update()
		}
	}
	function mainKeyboard{
		if ($_.KeyCode -eq "Enter"){
			if($NameBox.Focused){ $NameButton.PerformClick() }
			elseif ($PCBox.Focused){ $PCButton.PerformClick() }
			elseif ($IPBox.Focused){ $IPButton.PerformClick() }
			elseif ($UserIDBox.Focused){ $UserIDButton.PerformClick() }
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
		$NameBox.Text = ""
		$UserIDBox.Text = ""
		$PCBox.Text = ""
		$IPBox.Text = ""
		$LockoutBox.Text = ""
		$PhoneBox.Text = ""
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
		$IPBox.Forecolor = "black"
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
			$NameBox.Text = $global:Name
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
		$UserIDBox.Text = $global:UserID
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
		$PCBox.Text = $global:PCName
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
		if (($UserIDBox.Text).length -lt 1){
			$global:validID = $False
		}else{
		 #if($UserIDBox.Text -contains "@"){
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
		if (($PCBox.Text).length -lt 1){
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
			$IPBox.Forecolor = "red"
		}else{
			try{
				$IPObject = [System.Net.IPAddress]::parse($IP)
				[System.Net.IPAddress]::tryparse([string]$IP, [ref]$IPObject)
				([System.Net.Dns]::GetHostByAddress($global:IP)).HostName
			}catch{
				$global:IPWithARec = $False
			}
			$IPBox.Forecolor = "green"
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
		
		$UserIDBox.Text = $global:UserID
	}
	function PCNameByUserID{
		$variantid = "$UserID" + "*"
		$pcnames = @((Get-WmiObject -Class SMS_UserMachineRelationship -namespace "root\sms\$SCCMNameSpace" -computer $SCCMSiteServer -filter "UniqueUserName LIKE '%$global:UserID' and Types='1' and IsActive='1' and Sources='4'").ResourceName)
				
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
			$PCBox.Text = "-"
		}
			$PCBox.Text = $global:PCName

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
		$IPBox.Text = $global:IP
	}
	function LockoutByUserID{
		if ($global:UserID -ne "-"){
			$user = get-aduser $global:userID -properties LockedOut
			if ($user.LockedOut -eq $True){
				$LockoutBox.Text = "Locked"
				$LockoutButton.Visible = $True

			}else{
				$LockoutBox.Text = "Not Locked"
			}
		}else{
			$global:Lockout = "-"
			$LockoutBox.Text = "-"
		}
			
	}
	function PhoneByUserID{
		if ($global:UserID -ne "-"){
			$global:Phone = (Get-ADUser  $global:UserID -properties Pager).Pager #5 digit extension stored in pager property
		}else{
			$global:Phone = (Get-ADUser  $global:UserID -properties ipPhone).ipPhone
		}else{
			$global:Phone = (Get-ADUser  $global:UserID -properties OfficePhone).OfficePhone #full 10 digit or longer phone
		}else{
			$global:Phone = (Get-ADUser  $global:UserID -properties MobilePhone).MobilePhone
		}else{
			$global:Phone = "-"
		}
		
		$PhoneBox.Text = $global:Phone
		
		

	}
	function NameByUserID{
		if ($global:UserID -ne "-"){
			$global:Name = (Get-ADUser $global:UserID).Name
		}else{
			$global:Name = "-"
		}	

		$NameBox.Text = $global:Name

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
			}else{
				$returnval = $false
				MakeClickList ([ref]$global:UserID) $UserName ([ref]$returnval)
				if(!$returnval){$false}else{$true}
				
			}
		}
		$UserIDBox.Text = $global:UserID	
	}
	function PCNameByIP{
		$hostname = ([System.Net.Dns]::GetHostByAddress($IP)).HostName
		if ($hostname.indexof('.') -eq -1){
			$global:PCName = $hostname
		}else{
			$global:PCName = $hostname.Substring(0,$hostname.IndexOf('.'))
		}
		$PCBox.Text = $PCName

		
	}
	function HDriveByUserID{
		try{
			if ($global:UserID -ne "-"){
				$user = Get-AdUser $global:UserID -Properties HomeDirectory
				$HDriveBox.Text = $user.HomeDirectory
			}else{
				$HDriveBox.Text = "-"
			}
		}catch{
		
		}
		
	}
	function OUByUserID{
		try{
			if ($global:UserID -ne "-"){
				$user = Get-AdUser $global:UserID -Properties CanonicalName
				$OUBox.Text = $user.CanonicalName.Substring(($user.CanonicalName).IndexOf('/'))
			}else{
				$OUBox.Text = "-"
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
					$IPSourceLbl.Text ="(vpn)"
				}
			}
		}
	}
	function getADIP{
		try{
			$global:IP = (Get-ADComputer $global:PCName -Properties IPv4Address).IPv4Address
			$IPSourceLbl.Text ="(ad)"
		}catch{
			$global:IP = "-"
		}
	}
	function getDNSIP{
		$ips = @(([System.Net.Dns]::GetHostAddresses($global:PCName)).IPAddressToString)
		if ($ips.length -eq 1){
			$global:IP = $ips[0]
			$IPSourceLbl.Text ="(dns)"
			$true
		}elseif($ips.length -gt 1){
			$returnval = $false
			MakeClickList ([ref]$global:IP) $ips ([ref]$returnval)
			if(!$returnval){$false}else{$true}
		}
	}
	function pingIP{
		#$Online = Test-Connection -Computername $global:IP -BufferSize 16 -Count 1 -quiet
		#if($Online){ $IPBox.Forecolor = "green" } else{ $IPBox.Forecolor = "red" }
		$PingTimer.Enabled = $true
		$global:PingJob = Start-Job {
			param($IP)
			Test-Connection -Computername $IP -BufferSize 16 -Count 1 -quiet
		} -Arg $global:IP
	}
	function checkPing{
		if (($global:PingJob.State -eq "Completed")){
			$Pingable = @(Receive-Job -id $PingJob.id)
			if ($Pingable){
				$IPBox.Forecolor = "Green"
			}else{
				$IPBox.Forecolor = "Red"
			}
			$PingTimer.Enabled = $false
		}
	}
	
	### BUTTON CLICK FORMS ###
	function extraUserFacts{
		$user = Get-ADUser $UserID -properties LockedOut, Enabled, AccountExpirationDate, Certificates, Department, Description, PasswordNeverExpires, BadPwdCount, LastBadPasswordAttempt, PasswordLastSet, WhenChanged, WhenCreated
		
		$UserFactsForm = createForm "User Info" 300 350 "CenterScreen" "Fixed3D" $false $false $true
			#$UserFactsForm.Add_KeyDown({subWindowKeyListener($UserFactsForm)})
		$EnabledLbl = createItem "Label" 10 10 280 20 ("Enabled: "  + $user.Enabled) $UserFactsForm
		$AccountExpireLbl = createItem "Label" 10 40 280 20 ("Account Expires: " + $user.AccountExpirationDate) $UserFactsForm
		$PWNeverExpireLbl = createItem "Label" 10 70 280 20 ("Password Never Expires: " + $user.PasswordNeverExpires) $UserFactsForm
		$BadPWCountLbl = createItem "Label" 10 100 280 20 ("Bad Password Count: " + $user.BadPwdCount) $UserFactsForm
		$CreateTimeLbl = createItem "Label" 10 130 280 20 ("Created Time: " + $user.whenCreated) $UserFactsForm
		$ChangeTimeLbl = createItem "Label" 10 160 280 20 ("Modified Time: " + $user.whenChanged) $UserFactsForm
		$PWDaysLbl = createItem "Label" 10 190 280 20 (getPasswordAge) $UserFactsForm
		$CertBtn = createItem "Button" 2 217 10 20 "" $UserFactsForm
			$CertBtn.Visible = $False
			$global:CertStepper = 0
			$CertBtn.Add_Click({ stepCert })
		$CertLbl = createItem "Label" 10 220 270 20 (getCertText) $UserFactsForm
		$UserDepartLbl = createItem "Label" 10 250 280 20 ("Department: " + $user.Department) $UserFactsForm
		$UserDescriptLbl = createItem "Label" 10 280 280 20 ("Description: " + $user.Description) $UserFactsForm
		
		showForm $UserFactsForm
	}
	function extraPCFacts{
		$pc = Get-ADComputer $PCName -properties CanonicalName, Enabled, LastLogonDate, OperatingSystem, OperatingSystemServicePack, WhenChanged, WhenCreated
		
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
		$Enabled = createItem "Label" 10 10 270 20 ("Enabled: "  + $pc.Enabled) $PCFactsForm
		$OULbl = createItem "Label" 10 40 270 30 ("OU: " + ($pc.CanonicalName.Substring(($pc.CanonicalName).IndexOf('/')))) $PCFactsForm
		$LastLogonLbl = createItem "Label" 10 70 270 20 ("Last Logon: " + $pc.LastLogonDate) $PCFactsForm
		$OSLbl = createItem "Label" 10 100 270 20 ("OS: " + $pc.OperatingSystem + " " + $pc.OperatingSystemServicePack) $PCFactsForm
		$CreateTimeLbl = createItem "Label" 10 130 270 20 ("Created Time: " + $pc.whenCreated) $PCFactsForm
		$ChangeTimeLbl = createItem "Label" 10 160 270 20 ("Modified Time: " + $pc.whenChanged) $PCFactsForm
	
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
		$SoftVerButton = createItem "Button" 10 340 270 20 "Uptime and Software Versions" $PCFactsForm
			$SoftVerButton.Add_Click({ getSoftwareVersions })
		$SoftVerLbl = createItem "Label" 10 340 270 20 "Loading ... (May take a while)" $PCFactsForm
			$SoftVerLbl.ForeColor = "Blue"
			$SoftVerLbl.Visible = $false
		$UptimeLbl = createItem "Label" 10 340 300 20 "Uptime:" $PCFactsForm
		$IELbl = createItem "Label" 10 370 270 20 "IE:" $PCFactsForm
		$FlashLbl = createItem "Label" 10 400 270 20 "Flash:" $PCFactsForm
		$FlashButton = createItem "Button" 2 397 8 20 "" $PCFactsForm
			$FlashButton.Visible = $false
			$FlashButton.Add_Click({ stepFlash })
		$JavaLbl = createItem "Label" 10 430 270 20 "Java:" $PCFactsForm
		$JavaButton = createItem "Button" 2 427 8 20 "" $PCFactsForm
			$JavaButton.Visible = $false
			$JavaButton.Add_Click({ stepJava })
		#Endpoint Protection
		$EndPointButton = createItem "Button" 10 460 270 20 "View Endpoint Details" $PCFactsForm
			$EndPointButton.Add_Click({getEndpointInfo $EndPointButton $EndpointLabel})
		$EndPointLabel = createItem "Label" 10 460 270 20 "Signatures Last Updated: " $PCFactsForm
			$EndPointLabel.Visible = $False
		#View Smart Data
		$SmartLbl = createItem "Label" 100 490 270 20 "View Disk Health" $PCFactsForm
			$SmartLbl.ForeColor = "Blue"
			$SmartLbl.Add_Click({ getSmartData })

		showForm $PCFactsForm
	}
	function viewGroups($ADObjectType){
			$GroupForm = createForm "Groups" 250 300 "CenterScreen" "Fixed3D" $false $false $true
			$GroupList = createItem "ListBox" 10 10 210 180 "" $GroupForm
				$GroupList.HorizontalScrollbar = $true
				$GroupList.DataSource = @((getGroups($ADObjectType)).SamAccountName | Sort-Object )
			$AddGroupButton = createItem "Button" 10 190 210 30 "Add to Group" $GroupForm
				$AddGroupButton.Add_Click({ addGroup $ADObjectType})
			$RemoveGroupButton = createItem "Button" 10 225 210 30 "Remove from Group" $GroupForm
				$RemoveGroupButton.Add_Click({ removeGroup $ADObjectType})
			showForm($GroupForm)
	}
	function newUserPassword{
		$newPWForm = createForm "New Password" 300 160 "CenterScreen" "Fixed3D" $false $false $false
		$NewPWLbl = createItem "Label" 10 12 120 20 "Enter New Password: " $newPWForm
		$NewPWBox = createItem "TextBox" 140 10 120 20 "" $newPWForm
			$NewPWBox.PasswordChar = "*"
		$ConfirmPWLbl = createItem "Label" 10 42 120 20 "Confirm Password: " $newPWForm
		$ConfirmPWBox = createItem "TextBox" 140 40 120 20 "" $newPWForm
			$ConfirmPWBox.PasswordChar = "*"
		$ChangePWLoginCheck = createItem "CheckBox" 60 67 220 20 "Change password at next login" $newPWForm
		$NewPWOKButton = createItem "Button" 10 90 122 20 "OK" $newPWForm
		$NewPWCancelButton = createItem "Button" 147 90 122 20 "Cancel" $newPWForm

		$NewPWOKButton.Add_Click({
			if ($NewPWBox.Text -eq $ConfirmPWBox.text){
				try{
					$pw = ConvertTo-SecureString $NewPWBox.Text -AsPlainText -Force
					Set-ADAccountPassword $global:UserID -Reset -NewPassword $pw	
					if ($ChangePWLoginCheck.Checked -eq $true){
						Set-ADUser $global:UserID -ChangePasswordAtLogon:$true
					}
					$newPWForm.Close()
				}catch{
					$msg::Show($error[0])
				}
			}else{
				$msg::Show("Passwords do not match")
			}
		})
		
		$NewPWCancelButton.Add_Click({
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
			$CertBtn.Visible = $true
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
		$CertLbl.Text = listCert($index)
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
