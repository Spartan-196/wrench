$IconLocation = "$PSScriptRoot\icon.ico"
$LogoLocation = "$PSScriptRoot\logo.png"
$PSInfoLocation = "" #Path to psinfo.exe local or UNC https://docs.microsoft.com/en-us/sysinternals/downloads/psinfo 
$PSExecLocation = "" #Path to PsExec.exe local or UNC https://docs.microsoft.com/en-us/sysinternals/downloads/psexec 
$SmartCtlLocation = "" #Path to smartctl.exe local or UNC https://www.smartmontools.org
$SCCMRemoteLocation = "${env:ProgramFiles(x86)}\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe"  #Can be stored in UNC path if all required files to run are copied as well
$RenameComputerLocation = "" #point to prefered rename script file main ps1 is writen to launch .vbs or .vbe file
$VPNSiteUrl	= "" #Url on Cisco ASA for use with a username that has privlidge level 2 access to anyconnect connected users filter
				#Filer on ASA command line: show vpn-sessiondb detail {yourdb name} sort name | i Username|Assigned IP|
$ClientCenterLocation = "" #Path to SCCMCliCtrWPF.exe either local or UNC Share, utility from https://github.com/rzander/sccmclictr/releases
$SCCMSiteDataFile = "${env:ProgramFiles(x86)}\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager\ConfigurationManager.psd1"
$SCCMSiteServer = "" #Hostname for SCCM Primary
$SCCMNameSpace = "" #sms site name such as site_xx to buld WMI name space of root\sms\$sitename
$Domain = "" #Windows domain goes here so it can be excluded from resolved UserIDs in some functions
