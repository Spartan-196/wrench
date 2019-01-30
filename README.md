# WRENCH #

A project by [Matt Tuchfarber](https://github.com/tuchfarber)

Contributions by [James Atkinson](https://github.com/Spartan-196) and [Kevin Cook](https://github.com/PowershellWithPassion) 

Current Maintainer: [James Atkinson](https://github.com/Spartan-196)

## About ##

Wrench is a tool written in PowerShell to help minimize the number of tools needed by a PC Support team, while also providing all the information in a quickly viewable, easy application. Its primary tools for gathering information are the Active Directory tools in Powershell, and WMI calls to SCCM, and Windows clients.

## Running Wrench ##
Run the wrench.ps1 file, in my environment with separate user and admin accounts this had to be run in an elevated state to do this it a section of code by [Ben Armstrong on MSDN](https://blogs.msdn.microsoft.com/virtual_pc_guy/2010/09/23/a-self-elevating-powershell-script/) at the top of the was added to auto elevate the script to admin, this behavior can be commented out if you do not needed.
 

IF a team of people will be using this it can be stored on a network share and creating a shortcut to the wrench.bat file with an updated UNC in it. You can then place it on your desktop or add a custom toolbar to the taskbar to the folder on the network.


The computer Wrench is being ran on also must be running PowerShell 3.0 or greater, as some of the array functionality will break on Powershell 2. There is also a hotfix needed to allow unlocking of accounts on a Windows 7 PC, as the Powershell unlocking call is different than the standard Windows one. One or Both of these may be needed, they are KB2506143, [KB2577917](https://support.microsoft.com/en-us/help/2577917/unlocking-a-user-account-fails-when-using-adac-or-the-unlock-adaccount)

## The Main Screen ##

This is the main form of the application below will identify and explain all the elements on it. 

#### Four search bars ####
- Searching one item populates everything else on the form. For more information on searching, look for "Searching" below. 
- Requirements: None

<details><summary> Searching Details </summary>

###### If multiple items match Name, User name, or PC Name, a pop up list of items will appear, allowing you to select the correct one ######

###### Clicking the search button next to any field will only search that field and erase all the others until they are repopulated with new information ######

#### By Name ####
- Search in "Lastname, Firstname", "Firstname Lastname", "Lastname", "Firstname" format
- Matches to Name property of user in Active Directory
- Allows partial searches
- No character limits in search

- Finds (in order):
	1. ID by Name
	2. Phone by ID
	3. PC by ID
	4. IP by PC Name 
	5. Lockout by ID
	6. H Drive by ID
	7. OU by ID

#### By User ID ####
- Matches to SAMAccountName property of user in Active Directory
- Allows partial Searches
- No character limit in search 

- Finds (in order):
	1. Name by ID
	2. PC by ID
	3. IP by PC
	4. Lockout by ID
	5. Phone by ID
	6. H Drive by ID
	7. OU by ID

#### By PC Name ####
- Matches to Name property of computer in Active Directory
- Allows partial searches
- No character limits in search

- Finds (in order):
	1. ID by PC 
		- Finds user device affinity from SCCM and returns the user tied to PC
		- #1 *known bug of not getting SAM account of user returned*
	2. Name by ID
	3. IP by PC
	4. Lockout by ID
	5. Phone by ID
	6. H Drive by ID
	7. OU by ID

#### By IP Address ####
- Tests if IP address is valid IP address
- Attempts to get hostname of IP address using DNS 
- IP is in [+GREEN+] if the IP responds to Test-Connection, [-RED-] if it is not. **[-Red-] does not always mean offline**

- Finds (in order):
	1. PC by IP
	2. ID by PC
	3. Name by ID
	4. Lockout by ID
	5. Phone by ID
	6. H Drive by ID
	7. OU by ID
</Details>

#### Phone ####
- This is pulled from the Phone property of the user in Active Directory.
- In order it searches following AD propertry fields: Pager, ipPhone, OfficePhone, MobilePhone and returns the first one matched
- Requirements: User ID

#### Lockout ####
- If the user is unlocked it will say "Not Locked". 
- If the user is locked out, it will say "Locked" and a button will appear to unlock the account. 
- If it won't let you unlock the account, look into the hotfix noted in the "Running Wrench" section
- Requirements: User ID
- *Additional Requirements on Windows 7: KB2506143, KB2577917* 

#### H Drive ####
- Pulled from HomeDirectory property of user in Active Directory. 
- Requirements: User ID

#### User OU ####
- Pulled from CanonicalName property of user in Active Directory.
- Requirements: User ID

#### Connect Via SCCM Remote Control ####
- Uses the SCCM viewer found Defined in the wrench_env.ps1 file location for $SCCMRemoteLocation
- Connects using the IP address
- Requirements: IP Address

#### User details ####
- Checks if the user account is enabled (Enabled property)
- Checks when account is expired (AccountExpirationDate property)
- Checks if Password never expires (PasswordNeverExpires property)
- Checks bad password count (BadPwdCount property)
- Checks created time (whenCreated property)
- Checks modified time (whenChanged property)
- Checks password age (PasswordLastSet property)
- Checks issued certificates (Certificates property)
- Checks department (Department property)
- Checks description (Description property)
- Requirements: User ID

#### User Groups ####
- Displays groups that user is currently in, alphabetical sort.
- Add user to groups (See "Adding to Groups")
- Remove user from group (See "Removing from Groups")
- Requirements: User ID

#### Change Password ####
- Asks for new password, confirmed password (Set-ADAccountPassword) 
- Optional check box if it should be changed at next login (Set-ADUser -ChangePasswordAtLogon to true)
- Requirements: User ID, permissions to modify passwords

#### PC Details ####
- Checks if PC account is enabled (Enabled property)
- Checks OU of PC (CanonicalName property)
- Checks last date the computer was logged into (LastLoginDate property)
- Checks OS version (OperatingSystem and OperatingSystemServicePack property)
- Checks created time (whenCreated property)
- Checks modified time (whenChanged property)
- Can individually check:
	- Logged in user (WMI query)
	- PC type with Bios release date (WMI query)
	- MAC addresses of PC (getmac command)
	- Installed RAM (WMI) Displayed in GB with speed in MHz
	- Used, total, percent used of disk space (WMI win32_logicaldisk, Displayed in GB, with percentage used)
- Check Uptime, software versions of IE, Java, and Flash (Makes us of Sysinterals sysinfo and pulls selected information.)
- View Last update of Endpoint protection, or windows defender if running windows 10 (Remote Registry query)
- View Disk Health (PSexec of smartmontools)
- Requirements: PC Name, ***DNS must resolve correctly***

#### PC Groups ####
- Displays groups that computer is currently in, alphabetical sort.
- Adds computer to groups (See "Adding to Groups")
- Removes computer from groups (See "Removing from Groups")
- Requirements: PC Name


<details><summary> Adding to Groups </summary>
- To add a user or computer to a group, first click the (User/PC) Groups button
- Check if the (user/PC) is already in the group in the list box 
- If the (user/PC) is not in the list, click "Add to Group"
- A new pop up will appear allowing you to search for a group
	- Partial searches are enabled, though the unfiltered list is quite long
	- Search will be all groups that match the string entered ($string)
- After the list box populates, double click the item or press `ENTER` while it is selected to add to the group
- The form will exit after the addition has been made.

## Removing from a Group ##
- To remove a user or computer from a group, first click the (User/PC) Groups button
- Check that the (user/PC) is in the group you wish to remove it from first 
- If the (user/PC) is in the group, click "Remove (User/PC) from Group"
- Double click the name or press `ENTER` while it is selected to remove the (user/computer) from that group
- The form will exit after the removal has been made.
</Details>

#### Manage PC ####
- Pulls up Computer Managements window of remote PC
- If it can't connect, it will pull up Computer Management of current PC
- Requirements: PC Name, ***DNS must resolve correctly***

#### Rename PC ####
- Runs the Rename PC script currently written to run as a cscript for a vbs or vbe file  looking to make this more module.
- Requirements: Permissions to scripts directory, and to execute commands in said script
 

#### View C: ####
- Opens Windows Explorer to C$ of computer
	- Currently **does not work with Windows 10** using Microsoft SCM baseline policies.
	- Microsoft SCM baseline policy has `LocalAccountTokenFilterPolicy` set to `0` this effectivly disables being able to navigate to C$ of Windows 10 sysetems. A work around you must map a drive as "another user"
- Requirements: IP Address 

#### PS Remote ####
- Opens Remote PowerShell session to PC
- Requirements: PC Name ***DNS must resolve correctly***

#### RDP ####
- Remote Desktop into the PC
- Requirements: PC Name ***DNS must resolve correctly***

#### Keep Wrench on Top ####
- Will keep Wrench as the topmost item in Windows, covering up anything else.
- Useful if you wish to keep wrench docked on the side of the screen
	- **Warning**: If check box is checked and wrench is in the middle of the screen, it will cover the pop up windows in the program and may result in **unclickable** windows
- Requirements: None

## Expanding the Draw ##
- To expand the draw/side tray/lesser use buttons simply click the `>` button in the bottom right of the form
- Buttons available:
	- Check Group Policy
	- Telnet (See Telnet section)
	- Rename PC (See rename PC section)
	- SCCM Client Center (See Client Center section)

#### Telnet ####
- Opens telnet session to PC
- Requirements: Telnet client feature installed on support PC, IP Address

#### Client Center ####
- Launches [Client Center for Configuration Manager](https://github.com/rzander/sccmclictr) from the location defined by $ClientCenterLocation
- Used for running client side SCCM actions on a PC.
	- Most used toolbar buttons
		- SW Inv > Delta 
		- Machine Policy
		- User Policy 
		- App mgmt> Global and machine eval.
	- Commonly used area Software distribution and Inventory.
- Requirements: PC Name ***DNS must resolve correctly*** utilizes PSSession

