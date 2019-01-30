# WRENCH

A project by [Matt Tuchfarber](https://github.com/tuchfarber)

Contributions by [James Atkinson](https://github.com/Spartan-196) and [Kevin Cook](https://github.com/PowershellWithPassion) 

Current Maintainer: [James Atkinson](https://github.com/Spartan-196)

## About

Wrench is a tool written in PowerShell to help minimize the number of tools needed by a PC Support team, while also providing all the information in a quickly viewable, easy application. Its primary tools for gathering information are the Active Directory tools in Powershell, WMI calls to SCCM, and Windows clients.

## Running Wrench
Run the wrench.ps1 file. 

**Note**: In an environment with separate user and admin accounts, this has to be ran in an elevated state. This is done automatically by running a snippet of code by [Ben Armstrong on MSDN](https://blogs.msdn.microsoft.com/virtual_pc_guy/2010/09/23/a-self-elevating-powershell-script/) at application load. This behavior can be commented out if you do not need it.

If you're looking to deploy this to multiple users, we've found it easiest to keep their copy up to date by storing the `wrench.bat` file in a network location and creating a desktop shortcut to the file that can be optionally be added onto the taskbar. 

### Requirements
The computer Wrench is being ran on also must be running PowerShell 3.0 or greater, as some of the array functionality will break on Powershell 2. 

There is also a hotfix needed to allow unlocking of accounts on a Windows 7 PC, as the Powershell unlocking call is different than the standard Windows one. One or both of these may be needed, they are KB2506143, [KB2577917](https://support.microsoft.com/en-us/help/2577917/unlocking-a-user-account-fails-when-using-adac-or-the-unlock-adaccount)

## Using Wrench

### Searching

#### Details
* Searching one field populates the rest of the fields in the application
* If any search returns multiple values, a dialog will appear allowing you to select the correct one
* Searching any field immediately clears all existing data before pulling new information 
* Each field found will continue to traverse it's adjacent fields (per the relationships below) until all the data is populated.
![search field relationship model](docs/search_data_model.png?raw=true "Search field relationship model")

#### Editable Fields:
* **Name**
	- `Name` property of user in Active Directory
	- Search formats:
		- "Lastname, Firstname"
		- "Firstname Lastname"
		- "Lastname"
		- "Firstname"
	- Allows partial and no character limit searches
* **User ID**
	- `SAMAccountName` property of user in Active Directory
	- Allows partial and no character limit searches
* **PC Name**
	- `Name` property of computer in Active Directory
	- Allows partial and no character limit searches
* **IP Address**
	- Gets hostname of IP via DNS
	- IP will be [+GREEN+] if it responds to a `Test-Connection`, otherwise it will be [-RED-].
	- **[-Red-] does not always mean offline**

#### Read Only Fields:
* **Phone**
	- `Phone` property of the user in Active Directory.
	- Returns the first of the fields: `Pager`, `ipPhone`, `OfficePhone`, `MobilePhone`.
* **Lockout**
	- If the searched user account is locked, a button will appear to unlock the account. 
	- If it won't let you unlock the account, look into the hotfix noted in the "Running Wrench" section
* **H Drive**
	- `HomeDirectory` property of user in Active Directory. 
* **User OU**
	- `CanonicalName` property of user in Active Directory.
</Details>

### Actions

* **Connect Via SCCM Remote Control**
	- Connects to the IP address using the SCCM viewer found at `$SCCMRemoteLocation` as defined in `wrench_env.ps1`
	- Requires IP Address
* **Change Password**
	- Sets user's password to the password supplied in the prompt (`Set-ADAccountPassword`)
	- Can optionally set the password to be changed at next login via checkbox (`Set-ADUser -ChangePasswordAtLogon $TRUE`)
	- Requires UserID and permission to modify passwords
* **Manage PC**
	- Opens Computer Management (`compmgmt.msc`) window linked to remote PC
	- If connection fails, will open local PC's management 
	- Requires PC Name and ***correctly resolving DNS***
* **Rename PC**
	- Runs a renaming script which is currently written as a cscript for a vbs or vbe file
	- This script is not currently provided and therefore this functionality will not work out of the box
	- We are looking into make this a module
	- Requires permissions to the scripts directory and to execute the commands in the script
* **View C:**
	- Opens Exporer to `C$` of remote computer
	- Currently **does not work with Windows 10** using Microsoft SCM baseline policies.
	- Microsoft SCM baseline policy has `LocalAccountTokenFilterPolicy` set to `0` this effectivly disables being able to navigate to `C$` of Windows 10 systems. The work around being that you must map a drive as "another user"
	- Requires and IP address
* **PS Remote**
	- Opens Remote PowerShell session to the PC
	- Requires PC name and ***correctly resolving DNS***
* **RDP**
	- Remote Desktop into the PC
	- Requires PC Name and ***correctly resolving DNS***
* **Telnet**
	- Opens telnet session to PC
	- Requires IP address and telnet client feature installed on support PC
* **Client Center**
	- Launches [Client Center for Configuration Manager](https://github.com/rzander/sccmclictr) from the location found at `$ClientCenterLocation` as defined in `wrench_env.ps1`
	- Used for running client side SCCM actions on a PC.
		- Most used toolbar buttons
			- SW Inv > Delta 
			- Machine Policy
			- User Policy 
			- App mgmt> Global and machine eval.
		- Commonly used area Software distribution and Inventory.
	- Requires PC Name, PSSession, and ***correctly resolving DNS***

### Addtional Detail Dialogs

* **User details**
	<details>
	<summary> User details </summary>

	- Checks if the user account is enabled (`Enabled` property)
	- Checks when account is expired (`AccountExpirationDate` property)
	- Checks if Password never expires (`PasswordNeverExpires` property)
	- Checks bad password count (`BadPwdCount` property)
	- Checks created time (`whenCreated` property)
	- Checks modified time (`whenChanged` property)
	- Checks password age (`PasswordLastSet` property)
	- Checks issued certificates (`Certificates` property)
	- Checks department (`Department` property)
	- Checks description (`Description` property)
	- Requires User ID
	</details>

* **PC Details**
	<details>
	<summary> PC Details </summary>

	- Checks if PC account is enabled (`Enabled` property)
	- Checks OU of PC (`CanonicalName` property)
	- Checks last date the computer was logged into (`LastLoginDate` property)
	- Checks OS version (`OperatingSystem` and `OperatingSystemServicePack` property)
	- Checks created time (`whenCreated` property)
	- Checks modified time (`whenChanged` property)
	- Can individually check:
		- Logged in user (WMI query)
		- PC type with Bios release date (WMI query)
		- MAC addresses of PC (`getmac` command)
		- Installed RAM (WMI) Displayed in GB with speed in MHz
		- Used, total, percent used of disk space (WMI `win32_logicaldisk`, Displayed in GB, with percentage used)
	- Checks uptime and software versions of IE, Java, and Flash (Makes us of Sysinterals `sysinfo` and pulls selected information.)
	- View last update of Endpoint Protection, or Windows Defender if running Windows 10 (Remote Registry query)
	- View Disk Health (PSexec of smartmontools)
	- Requires PC Name and ***correctly resolving DNS***
	</details>

### Groups 

#### Type of Groups
* **User Groups**
	- Displays groups that user is currently in in alphabetical order.
	- Can add [add](#Adding-to-Groups) or [remove](#Removing-from-Groups) user from groups
	- Requires User ID
* **PC Groups**
	- Displays groups that computer is currently in in alphabetical order.
	- Can add [add](#Adding-to-Groups) or [remove](#Removing-from-Groups) computer from groups
	- Requires PC Name


#### Adding to Groups
1. To add a user or computer to a group, first click the (user/PC) Groups button
2. Verify the (user/PC) isn't already in the group via the list 
3. If the (user/PC) isn't in the list, click "Add to Group"
4. A new dialog will appear allowing you to search for a group
	- Partial searches are enabled, though the unfiltered list can be quite long
5. After searching, double click the group or highlight and press `ENTER` to add the (user/PC) to the group
6. The dialog will close after the addition has been made.

#### Removing from Groups
1. To remove a user or computer from a group, first click the (User/PC) Groups button
2. Verify the (user/PC) is in the group you wish to remove it from
3. If the (user/PC) is in the group, click "Remove (User/PC) from Group"
4. double click the group or highlight and press `ENTER` to remove the (user/computer) from that group
- The dialog will close after the removal has been made.


### Additional Features
* **Keep Wrench on Top**
	- Will keep Wrench as the topmost item in Windows, covering up anything else.
	- Useful if you wish to keep wrench docked on the side of the screen
	- **Warning**: If checked and wrench is located in the middle of the screen, it will cover the pop up dialogs and may result in **unclickable** windows
* **Expanding the Drawer**
	- The bottom right corner has a `>` button which, if clicked, will expand to show lesser used buttons
	- Buttons currently available:
		- Check Group Policy
		- Telnet (See Telnet section)
		- Rename PC (See rename PC section)
		- SCCM Client Center (See Client Center section)
