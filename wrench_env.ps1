<?xml version="1.0"?>
<configuration>
		<appSettings>
		<!--Vars -->
			<add key="IconLocation" value=".\icon.ico"/>
			<add key="LogoLocation" value=".\logo.png"/>
			<add key ="PSInfoLocation" value=""/>
			<add key="PSExecLocation" value=""/>
			<add key="SmartCtlLocation" value=""/>
			<add key="SCCMRemoteLocation" value="ProgramFiles(x86)\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe"/>
			<add key="ClientCenterLocation" value=""/>
			<add key="SCCMSiteDataFile" value="ProgramFiles(x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager\ConfigurationManager.psd1"/>
			<add key="RenameComputerLocation" value=""/>
			<add key="VPNSiteUrl" value=""/>
			<add key="SCCMSiteServer" value=""/>
			<add key="SCCMSiteCode" value=""/>
			<add key="SCCMNameSpace" value=""/>		
			<add key="Domain" value=""/>
			<add key="OpenForUse" value=""/>
		</appSettings>
</configuration>

<!--
#$IconLocation = "$PSScriptRoot\icon.ico"
#$LogoLocation = "$PSScriptRoot\logo.png"
#$PSInfoLocation = "" #Path to psinfo.exe local or UNC https://docs.microsoft.com/en-us/sysinternals/downloads/psinfo 
#$PSExecLocation = "" #Path to PsExec.exe local or UNC https://docs.microsoft.com/en-us/sysinternals/downloads/psexec 
#$SmartCtlLocation = "" #Path to smartctl.exe local or UNC https://www.smartmontools.org
#$SCCMRemoteLocation = "${env:ProgramFiles(x86)}\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe"  #Can be stored in UNC path if all required files to run are copied as well
#$RenameComputerLocation = "" #point to prefered rename script file main ps1 is writen to launch .vbs or .vbe file
#$VPNSiteUrl	= "" #Url on Cisco ASA for use with a username that has privlidge level 2 access to anyconnect connected users filter
				#Filer on ASA command line: show vpn-sessiondb detail {yourdb name} sort name | i Username|Assigned IP|
#$ClientCenterLocation = "" #Path to SCCMCliCtrWPF.exe either local or UNC Share, utility from https://github.com/rzander/sccmclictr/releases
#$SCCMSiteDataFile = "${env:ProgramFiles(x86)}\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager\ConfigurationManager.psd1"
#$SCCMSiteServer = "" #Hostname for SCCM Primary
# The $SCCMNameSpace variable is obsolete and has been removed from the code.
#$SCCMNameSpace = "" #sms site name such as site_xx to buld WMI name space of root\sms\$sitename 
#$Domain = "" #Windows domain goes here so it can be excluded from resolved UserIDs in some functions
#$SCCMSiteCode = "" # Add three letter site code. ex. $SCCMSiteCode = "ABC". You can retrive the site code with following code from a machine the Wrench works on. ([wmiclass]"ROOT\ccm:SMS_Client").GetAssignedSite().sSiteCode
-->