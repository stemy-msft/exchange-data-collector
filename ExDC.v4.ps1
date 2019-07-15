<#
.SYNOPSIS
	Collects data in a Microsoft Exchange environment and assembles the output into Word and Excel files.
.DESCRIPTION
	ExDC is used to collect a large amount of information about an Exchange environment with minimal effort.
	The data is initially collected into a series of text files that can then be assembled into reports on the
	data collection server or another workstation.  This script was originally written for use by
	Microsoft Premier Services engineers during onsite engagements.

	Script guidelines:
	* Complete data collection requires elevated credentials for both Exchange and Active Directory
	* It is recommended that ExDC not run directly on Exchange server with a production load
	* Data collection does not require that Office is installed as output files are txt and xml
	* The ExDC folder can be forklifted to a workstation with Office to generate output reports
.PARAMETER JobCount_ExOrg
	Max number of jobs for Exchange cmdlet functions (Default = 10)
	Caution: The OOB throttling policy sets PowershellMaxConcurrency at 18 sessions per user per server
	Modifying this value without increasing the throttling policy can cause ExOrg jobs to immediately fail.
.PARAMETER JobCount_WMI
	Max number of jobs for non-Exchange cmdlet functions (Default = 25)
.PARAMETER JobPolling_ExOrg
	Polling interval for job completion for Exchange cmdlet functions (Default = 5 sec)
.PARAMETER JobPolling_WMI
	Polling interval for job completion for non-Exchange functions  (Default = 5 sec)
.PARAMETER Timeout_ExOrg_Job
	Job timeout for Exchange functions  (Default = 3600 sec)
	The default value is 3600 seconds but should be adjusted for organizations with a large number of mailboxes or servers over slow connections.
.PARAMETER Timeout_WMI_Job
	Job timeout for non-Exchange functions  (Default = 600 sec)
	If a job exceeds this value, it is terminated at the next interval specified in JobPolling_WMI.
	The default value is 600 seconds but should be adjusted for organizations with servers over slow connections.
.PARAMETER ServerForPSSession
	Exchange server used for Powershell sessions
.PARAMETER INI_Server
	Specify INI file for Server Tests configuration
.PARAMETER INI_Cluster
	Specify INI file for Cluster Tests configuration
.PARAMETER INI_ExOrg
	Specify INI file for ExOrg Tests configuration
.PARAMETER NoEMS
	Use this switch to launch the tool in Powershell (No Exchange cmdlets)
.PARAMETER NoGUI
	Use this switch with template files specified automatically start data collection.
	Any additional files required (mailbox.txt, dc.txt, and/or exchange.txt) must be present.
	The -ServerForPSSession parameter must also be specified
	This is useful for scheduling the data collection using Task Scheduler.
.PARAMETER NoGUIOutputFolder
	This switch is only used when the -NoGUI switch is used.
	When specified, the output folder will be renamed to this value after data collection completes
	If not specified, the output folder will be rename to "output-x" where x is the Get-Time tick value
.EXAMPLE
	.\ExDC.v4.ps1 -JobCount_ExOrg 12
	This results in ExDC using 12 active ExOrg jobs instead of the default of 10.
.EXAMPLE
	.\ExDC.v4.ps1 -JobPolling_ExOrg 30
	This results in ExDC polling for completed ExOrg jobs every 30 seconds.
.EXAMPLE
	.\ExDC.v4.ps1 -Timeout_ExOrg_Job 7200
	This results in ExDC killing ExOrg jobs that have exceeded 7200 seconds at the next polling interval.
.EXAMPLE
	.\ExDC.v4.ps1 -INI_Server ".\Templates\Template_Recommended_INI_Server.ini"
	This results in ExDC loading the specified template on start up.
.EXAMPLE
	.\ExDC.v4.ps1 -NoGUI -INI_Server ".\Templates\Template_Recommended_INI_Server.ini" -NoGUIOutputFolder "Output-Completed" -ServerForPSSession "Exchange2016.domain.com"
	This results in ExDC running against Exchange2016.domain.com without the GUI and executing the tests indicated in the INI file
	When completed, ExDC will rename the output folder in the ExDC location to "Output-Completed"
.INPUTS
	None.
.OUTPUTS
	This script has no output objects.  ExDC creates txt, xml, docx, and xlsx output.
.NOTES
	NAME        :   ExDC.v4.ps1
	AUTHOR      :   Stemy Mynhier [MSFT]
	VERSION     :   4.0.2 build a1
	LAST EDIT   :   Oct-2018
.LINK
	https://gallery.technet.microsoft.com/office/Exchange-Data-Collector-ed48c3db
	https://github.com/stemy-msft/exchange-data-collector
#>

Param(	[int]$JobCount_ExOrg = 10,`
		[int]$JobPolling_ExOrg = 5,`
		[int]$JobCount_WMI = 25,`
		[int]$JobPolling_WMI = 5,`
		[int]$Timeout_WMI_Job = 600,`
		[int]$Timeout_ExOrg_Job = 3600,`
		[string]$ServerForPSSession = $null,`
		[string]$INI_Server,`
		[string]$INI_Cluster,`
		[string]$INI_ExOrg,`
		[switch]$NoEMS,`
		[switch]$NoGUI,`
		[string]$NoGUIOutputFolder)

function New-ExDCForm {
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null

#region *** Initialize Form ***

#region Main Form
$form1 = New-Object System.Windows.Forms.Form
$tab_Master = New-Object System.Windows.Forms.TabControl
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
$Menu_Main = new-object System.Windows.Forms.MenuStrip
$Menu_File = new-object System.Windows.Forms.ToolStripMenuItem('&File')
$Menu_Toggle = new-object System.Windows.Forms.ToolStripMenuItem('&Toggle')
$Menu_Help  = new-object System.Windows.Forms.ToolStripMenuItem('&Help')
$Submenu_LoadTargets = new-object System.Windows.Forms.ToolStripMenuItem('&Load all Targets from files')
$Submenu_PackageLogs = new-object System.Windows.Forms.ToolStripMenuItem('&Package application log')
$Submenu_Targets_CheckAll = new-object System.Windows.Forms.ToolStripMenuItem('&Check All Targets')
$Submenu_Targets_UnCheckAll = new-object System.Windows.Forms.ToolStripMenuItem('&Uncheck All Targets')
$Submenu_Tests_CheckAll = new-object System.Windows.Forms.ToolStripMenuItem('&Check All Tests')
$Submenu_Tests_UnCheckAll = new-object System.Windows.Forms.ToolStripMenuItem('&Uncheck All Tests')
$Submenu_Help = new-object System.Windows.Forms.ToolStripMenuItem('&Help')
$Submenu_About = new-object System.Windows.Forms.ToolStripMenuItem('&About')
#endregion Main Form

#region Step1 - Targets

#region Step1 Main
$tab_Step1 = New-Object System.Windows.Forms.TabPage
$btn_Step1_Discover = New-Object System.Windows.Forms.Button
$btn_Step1_Populate = New-Object System.Windows.Forms.Button
$tab_Step1_Master = New-Object System.Windows.Forms.TabControl
$status_Step1 = New-Object System.Windows.Forms.StatusBar
#endregion Step1 Main

#region Step1 DC Tab
$tab_Step1_DC = New-Object System.Windows.Forms.TabPage
$bx_DC_List = New-Object System.Windows.Forms.GroupBox
$btn_Step1_DC_Discover = New-Object System.Windows.Forms.Button
$btn_Step1_DC_Populate = New-Object System.Windows.Forms.Button
$clb_Step1_DC_List = New-Object system.Windows.Forms.CheckedListBox
$txt_DCTotal = New-Object System.Windows.Forms.TextBox
$btn_Step1_DC_CheckAll = New-Object System.Windows.Forms.Button
$btn_Step1_DC_UncheckAll = New-Object System.Windows.Forms.Button
#endregion Step1 DC Tab

#region Step1 Exchange Tab
$tab_Step1_Ex = New-Object System.Windows.Forms.TabPage
$bx_Ex_List = New-Object System.Windows.Forms.GroupBox
$btn_Step1_Ex_Discover = New-Object System.Windows.Forms.Button
$btn_Step1_Ex_Populate = New-Object System.Windows.Forms.Button
$clb_Step1_Ex_List = New-Object system.Windows.Forms.CheckedListBox
$txt_ExchTotal = New-Object System.Windows.Forms.TextBox
$btn_Step1_Ex_CheckAll = New-Object System.Windows.Forms.Button
$btn_Step1_Ex_UncheckAll = New-Object System.Windows.Forms.Button
#endregion Step1 Exchange Tab

#region Step1 Nodes Tab
if (($Exchange2007Powershell -eq $true) -or ($Exchange2003orEarlier -eq $true) -or ($NoEMS -eq $true) -or ($NoGUI -eq $true))
{
	$tab_Step1_Nodes = New-Object System.Windows.Forms.TabPage
	$bx_Nodes_List = New-Object System.Windows.Forms.GroupBox
	$btn_Step1_Nodes_Discover = New-Object System.Windows.Forms.Button
	$btn_Step1_Nodes_Populate = New-Object System.Windows.Forms.Button
	$clb_Step1_Nodes_List = New-Object system.Windows.Forms.CheckedListBox
	$txt_NodesTotal = New-Object System.Windows.Forms.TextBox
	$btn_Step1_Nodes_CheckAll = New-Object System.Windows.Forms.Button
	$btn_Step1_Nodes_UncheckAll = New-Object System.Windows.Forms.Button
}
#endregion Step1 Nodes Tab

#region Step1 Mailboxes Tab
if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true) -or ($NoGUI -eq $true))
{
	$tab_Step1_Mailboxes = New-Object System.Windows.Forms.TabPage
	$bx_Mailboxes_List = New-Object System.Windows.Forms.GroupBox
	$btn_Step1_Mailboxes_Discover = New-Object System.Windows.Forms.Button
	$btn_Step1_Mailboxes_Populate = New-Object System.Windows.Forms.Button
	$clb_Step1_Mailboxes_List = New-Object system.Windows.Forms.CheckedListBox
	$txt_MailboxesTotal = New-Object System.Windows.Forms.TextBox
	$btn_Step1_Mailboxes_CheckAll = New-Object System.Windows.Forms.Button
	$btn_Step1_Mailboxes_UncheckAll = New-Object System.Windows.Forms.Button
}
#endregion Step1 Mailboxes Tab

#endregion Step1 - Targets

#region Step2 - Templates
$tab_Step2 = New-Object System.Windows.Forms.TabPage
$bx_Step2_Templates = New-Object System.Windows.Forms.GroupBox
$rb_Step2_Template_1 = New-Object System.Windows.Forms.RadioButton
$rb_Step2_Template_2 = New-Object System.Windows.Forms.RadioButton
$rb_Step2_Template_3 = New-Object System.Windows.Forms.RadioButton
$rb_Step2_Template_4 = New-Object System.Windows.Forms.RadioButton
$Status_Step2 = New-Object System.Windows.Forms.StatusBar
#endregion Step2 - Templates


#Region Step3 - Tests

#region Step3 Main Tier1
$tab_Step3 = New-Object System.Windows.Forms.TabPage
$tab_Step3_Master = New-Object System.Windows.Forms.TabControl
$status_Step3 = New-Object System.Windows.Forms.StatusBar
$lbl_Step3_Execute = New-Object System.Windows.Forms.Label
$btn_Step3_Execute = New-Object System.Windows.Forms.Button
#endregion Step3 Main Tier1

#region Step3 Server Tier2
$tab_Step3_Server = New-Object System.Windows.Forms.TabPage
$tab_Step3_Server_Tier2 = New-Object System.Windows.Forms.TabControl
#endregion Step3 Server Tier2

#region Step3 ExOrg Tier2
if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true) -or ($NoGUI -eq $true))
{
	$tab_Step3_ExOrg = New-Object System.Windows.Forms.TabPage
	$tab_Step3_ExOrg_Tier2 = New-Object System.Windows.Forms.TabControl
}
#endregion Step3 ExOrg Tier2

#region Step3 DC Tab
$tab_Step3_DC = New-Object System.Windows.Forms.TabPage
$bx_DC_Functions = New-Object System.Windows.Forms.GroupBox
$chk_DC_Win32_Bios = New-Object System.Windows.Forms.CheckBox
$chk_DC_Win32_ComputerSystem = New-Object System.Windows.Forms.CheckBox
$chk_DC_Win32_LogicalDisk = New-Object System.Windows.Forms.CheckBox
$chk_DC_Win32_NetworkAdapter = New-Object System.Windows.Forms.CheckBox
$chk_DC_Win32_NetworkAdapterConfig = New-Object System.Windows.Forms.CheckBox
$chk_DC_Win32_OperatingSystem = New-Object System.Windows.Forms.CheckBox
$chk_DC_Win32_PageFileUsage = New-Object System.Windows.Forms.CheckBox
$chk_DC_Win32_PhysicalMemory = New-Object System.Windows.Forms.CheckBox
$chk_DC_Win32_Processor = New-Object System.Windows.Forms.CheckBox
$chk_DC_Registry_AD = New-Object System.Windows.Forms.CheckBox
$chk_DC_Registry_OS = New-Object System.Windows.Forms.CheckBox
$chk_DC_Registry_Software = New-Object System.Windows.Forms.CheckBox
$chk_DC_MicrosoftDNS_Zone = New-Object System.Windows.Forms.CheckBox
$chk_DC_MSAD_DomainController = New-Object System.Windows.Forms.CheckBox
$chk_DC_MSAD_ReplNeighbor = New-Object System.Windows.Forms.CheckBox
$btn_Step3_DC_CheckAll = New-Object System.Windows.Forms.Button
$btn_Step3_DC_UncheckAll = New-Object System.Windows.Forms.Button
#endregion Step3 DC Tab

#region Step3 Exchange Tab
$tab_Step3_Exchange = New-Object System.Windows.Forms.TabPage
$bx_Exchange_Functions = New-Object System.Windows.Forms.GroupBox
$chk_Ex_Win32_Bios = New-Object System.Windows.Forms.CheckBox
$chk_Ex_Win32_ComputerSystem = New-Object System.Windows.Forms.CheckBox
$chk_Ex_Win32_LogicalDisk = New-Object System.Windows.Forms.CheckBox
$chk_Ex_Win32_NetworkAdapter = New-Object System.Windows.Forms.CheckBox
$chk_Ex_Win32_NetworkAdapterConfig = New-Object System.Windows.Forms.CheckBox
$chk_Ex_Win32_OperatingSystem = New-Object System.Windows.Forms.CheckBox
$chk_Ex_Win32_PageFileUsage = New-Object System.Windows.Forms.CheckBox
$chk_Ex_Win32_PhysicalMemory = New-Object System.Windows.Forms.CheckBox
$chk_Ex_Win32_Processor = New-Object System.Windows.Forms.CheckBox
$chk_Ex_Registry_Ex = New-Object System.Windows.Forms.CheckBox
$chk_Ex_Registry_OS = New-Object System.Windows.Forms.CheckBox
$chk_Ex_Registry_Software = New-Object System.Windows.Forms.CheckBox
$btn_Step3_Ex_CheckAll = New-Object System.Windows.Forms.Button
$btn_Step3_Ex_UncheckAll = New-Object System.Windows.Forms.Button
#endregion Step3 Exchange Tab

#region Step3 Cluster Tab
if (($Exchange2007Powershell -eq $true) -or ($Exchange2003orEarlier -eq $true) -or ($NoEMS -eq $true) -or ($NoGUI -eq $true))
{
	$tab_Step3_Cluster = New-Object System.Windows.Forms.TabPage
	$bx_Cluster_Functions = New-Object System.Windows.Forms.GroupBox
	$chk_Cluster_MSCluster_Node = New-Object System.Windows.Forms.CheckBox
	$chk_Cluster_MSCluster_Network = New-Object System.Windows.Forms.CheckBox
	$chk_Cluster_MSCluster_Resource = New-Object System.Windows.Forms.CheckBox
	$chk_Cluster_MSCluster_ResourceGroup = New-Object System.Windows.Forms.CheckBox
	$btn_Step3_Cluster_CheckAll = New-Object System.Windows.Forms.Button
	$btn_Step3_Cluster_UncheckAll = New-Object System.Windows.Forms.Button
}
#endregion Step3 Cluster Tab

#region Step3 Client Access tab
if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true) -or ($NoGUI -eq $true))
{
	$tab_Step3_ClientAccess = New-Object System.Windows.Forms.TabPage
	$bx_ClientAccess_Functions = New-Object System.Windows.Forms.GroupBox
	$btn_Step3_ClientAccess_CheckAll = New-Object System.Windows.Forms.Button
	$btn_Step3_ClientAccess_UncheckAll = New-Object System.Windows.Forms.Button
	$chk_Org_Get_ActiveSyncDevice = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_ActiveSyncPolicy = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_ActiveSyncVirtualDirectory = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_AutodiscoverVirtualDirectory = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_AvailabilityAddressSpace = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_ClientAccessArray = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_ClientAccessServer = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_ECPVirtualDirectory = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_OABVirtualDirectory = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_OutlookAnywhere = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_OWAMailboxPolicy = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_OWAVirtualDirectory = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_PowershellVirtualDirectory = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_RPCClientAccess = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_ThrottlingPolicy = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_WebServicesVirtualDirectory = New-Object System.Windows.Forms.CheckBox
}

#endregion Step3 Client Access tab

#region Step3 Global tab
if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true) -or ($NoGUI -eq $true))
{
	$tab_Step3_Global = New-Object System.Windows.Forms.TabPage
	$bx_Global_Functions = New-Object System.Windows.Forms.GroupBox
	$btn_Step3_Global_CheckAll = New-Object System.Windows.Forms.Button
	$btn_Step3_Global_UncheckAll = New-Object System.Windows.Forms.Button
	$chk_Org_Get_AddressBookPolicy  = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_AddressList  = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_DatabaseAvailabilityGroup  = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_DAGNetwork  = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_EmailAddressPolicy = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_ExchangeCertificate = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_ExchangeServer = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_MailboxDatabase = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_MailboxDatabaseCopyStatus = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_MailboxServer = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_OfflineAddressBook = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_OrgConfig = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_PublicFolderDatabase = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_Rbac = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_RetentionPolicy = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_RetentionPolicyTag = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_StorageGroup = New-Object System.Windows.Forms.CheckBox
}
#endregion Step3 Global tab

#region Step3 Recipient Tab
if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true) -or ($NoGUI -eq $true))
{
	$tab_Step3_Recipient = New-Object System.Windows.Forms.TabPage
	$bx_Recipient_Functions = New-Object System.Windows.Forms.GroupBox
	$btn_Step3_Recipient_CheckAll = New-Object System.Windows.Forms.Button
	$btn_Step3_Recipient_UncheckAll = New-Object System.Windows.Forms.Button
	$chk_Org_Get_ADPermission = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_CalendarProcessing = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_CASMailbox = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_DistributionGroup = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_DynamicDistributionGroup = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_Mailbox = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_MailboxFolderStatistics = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_MailboxPermission = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_MailboxStatistics = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_PublicFolder = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_PublicFolderStatistics = New-Object System.Windows.Forms.CheckBox
    $chk_Org_Get_User = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Quota = New-Object System.Windows.Forms.CheckBox
}
#endregion Step3 Recipient Tab

#region Step3 Transport Tab
if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true) -or ($NoGUI -eq $true))
{
	$tab_Step3_Transport = New-Object System.Windows.Forms.TabPage
	$bx_Transport_Functions = New-Object System.Windows.Forms.GroupBox
	$btn_Step3_Transport_CheckAll = New-Object System.Windows.Forms.Button
	$btn_Step3_Transport_UncheckAll = New-Object System.Windows.Forms.Button
	$chk_Org_Get_AcceptedDomain = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_AdSite = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_AdSiteLink = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_ContentFilterConfig = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_ReceiveConnector = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_RemoteDomain = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_RoutingGroupConnector = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_SendConnector = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_TransportConfig = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_TransportRule = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_TransportServer = New-Object System.Windows.Forms.CheckBox
}
#endregion Step3 Transport Tab

#region Step3 Unified Messaging tab
if ((($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true) -or ($NoGUI -eq $true)) -and ($UM -eq $true))
{
	$tab_Step3_UM = New-Object System.Windows.Forms.TabPage
	$bx_UM_Functions = New-Object System.Windows.Forms.GroupBox
	$btn_Step3_UM_CheckAll = New-Object System.Windows.Forms.Button
	$btn_Step3_UM_UncheckAll = New-Object System.Windows.Forms.Button
	$chk_Org_Get_UmAutoAttendant = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_UmDialPlan = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_UmIpGateway = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_UmMailbox = New-Object System.Windows.Forms.CheckBox
	#$chk_Org_Get_UmMailboxConfiguration = New-Object System.Windows.Forms.CheckBox
	#$chk_Org_Get_UmMailboxPin = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_UmMailboxPolicy = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_UmServer = New-Object System.Windows.Forms.CheckBox
}
#endregion Step3 Unified Messaging tab

#region Step3 Misc Tab
if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true) -or ($NoGUI -eq $true))
{
	$tab_Step3_Misc = New-Object System.Windows.Forms.TabPage
	$bx_Misc_Functions = New-Object System.Windows.Forms.GroupBox
	$btn_Step3_Misc_CheckAll = New-Object System.Windows.Forms.Button
	$btn_Step3_Misc_UncheckAll = New-Object System.Windows.Forms.Button
	$chk_Org_Get_AdminGroups = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_Fsmo = New-Object System.Windows.Forms.CheckBox
	$chk_Org_Get_ExchangeServerBuilds = New-Object System.Windows.Forms.CheckBox
}
#endregion Step3 Misc Tab

#EndRegion Step3 - Tests

#region Step4 - Reporting
$tab_Step4 = New-Object System.Windows.Forms.TabPage
$btn_Step4_Assemble = New-Object System.Windows.Forms.Button
$lbl_Step4_Assemble = New-Object System.Windows.Forms.Label
$bx_Step4_Functions = New-Object System.Windows.Forms.GroupBox
$chk_Step4_DC_Report = New-Object System.Windows.Forms.CheckBox
$chk_Step4_Ex_Report = New-Object System.Windows.Forms.CheckBox
$chk_Step4_ExOrg_Report = New-Object System.Windows.Forms.CheckBox
$chk_Step4_Exchange_Environment_Doc = New-Object System.Windows.Forms.CheckBox
$Status_Step4 = New-Object System.Windows.Forms.StatusBar
#endregion Step4 - Reporting

#region Step5 - Having Trouble?
#$tab_Step5 = New-Object System.Windows.Forms.TabPage
#$bx_Step5_Functions = New-Object System.Windows.Forms.GroupBox
#$Status_Step5 = New-Object System.Windows.Forms.StatusBar
#endregion Step5 - Having Trouble?

#endregion *** Initialize Form ***

#region *** Events ***

#region "Main Menu" Events

$handler_Submenu_LoadTargets=
{
	Import-TargetsDc
	Import-TargetsEx
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2003orEarlier -eq $true) -or ($NoEMS -eq $true) -or ($NoGUI -eq $true))
	{
		Import-TargetsNodes
	}
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true) -or ($NoGUI -eq $true))
	{
		Import-TargetsMailboxes
	}
}

$handler_Submenu_PackageLogs=
{
	.\ExDC_Scripts\Core_Package_Logs.ps1 -RunLocation $location
}

$handler_Submenu_Targets_CheckAll=
{
	Enable-TargetsDc
	Enable-TargetsEx
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2003orEarlier -eq $true) -or ($NoEMS -eq $true) -or ($NoGUI -eq $true))
	{
		Enable-TargetsNodes
	}
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true) -or ($NoGUI -eq $true))
	{
		Enable-TargetsMailbox
	}
}

$handler_Submenu_Targets_UnCheckAll=
{
	Disable-TargetsDc
	Disable-TargetsEx

	if (($Exchange2007Powershell -eq $true) -or ($Exchange2003orEarlier -eq $true) -or ($NoEMS -eq $true) -or ($NoGUI -eq $true))
	{
		Disable-TargetsNodes
	}
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true) -or ($NoGUI -eq $true))
	{
		Disable-TargetsMailbox
	}
}

$handler_Submenu_Tests_CheckAll=
{
	# Server Functions - Domain Controllers
	Set-AllFunctionsDc -check $true
	# Server Functions - Exchange Servers
	Set-AllFunctionsEx -check $true
	# Server Functions - Cluster Nodes
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2003orEarlier -eq $true) -or ($NoEMS -eq $true) -or ($NoGUI -eq $true))
	{
		Set-AllFunctionsCluster -Check $true
	}
	# Exchange Functions - All
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true) -or ($NoGUI -eq $true))
	{
		Set-AllFunctionsClientAccess -Check $true
		Set-AllFunctionsGlobal -Check $true
		Set-AllFunctionsRecipient -Check $true
		Set-AllFunctionsTransport -Check $true
		Set-AllFunctionsMisc -Check $true
	}
	# UM is special
	if ((($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true) -or ($NoGUI -eq $true)) -and ($UM -eq $true))
	{
		Set-AllFunctionsUm -Check $true
	}
}

$handler_Submenu_Tests_UnCheckAll=
{
	# Server Functions - Domain Controllers
	Set-AllFunctionsDc -Check $False
	# Server Functions - Exchange Servers
	Set-AllFunctionsEx -Check $false
	# Server Functions - Cluster Nodes
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2003orEarlier -eq $true) -or ($NoEMS -eq $true) -or ($NoGUI -eq $true))
	{
		Set-AllFunctionsCluster -Check $False
	}
	# Exchange Functions - All
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true) -or ($NoGUI -eq $true))
	{
		Set-AllFunctionsClientAccess -Check $False
		Set-AllFunctionsGlobal -Check $False
		Set-AllFunctionsRecipient -Check $False
		Set-AllFunctionsTransport -Check $False
		Set-AllFunctionsMisc -Check $False

	}
	# UM is special
	if ((($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true) -or ($NoGUI -eq $true)) -and ($UM -eq $true))
	{
		Set-AllFunctionsUm -Check $False
	}
}

$handler_Submenu_Help=
{
	$Message_Help = "Would you like to open the Help document?"
	$Title_Help = "ExDC Help"
	$MessageBox_Help = [Windows.Forms.MessageBox]::Show($Message_Help, $Title_Help, [Windows.Forms.MessageBoxButtons]::YesNo, [Windows.Forms.MessageBoxIcon]::Information)
	if ($MessageBox_Help -eq [Windows.Forms.DialogResult]::Yes)
	{
		$ie = New-Object -ComObject "InternetExplorer.Application"
		$ie.visible = $true
		$ie.navigate((get-location).path + "\Help\Documentation_ExDC.v.4.mht")
	}
}

$handler_Submenu_About=
{
	$Message_About = ""
	$Message_About = "Exchange Data Collector `n`n"
	$Message_About = $Message_About += "Version: 4.0.2 Build a1 `n`n"
	$Message_About = $Message_About += "Release Date: September 2018 `n`n"
	$Message_About = $Message_About += "Written by: Stemy Mynhier`nstemy@microsoft.com `n`n"
	$Message_About = $Message_About += "This script is provided AS IS with no warranties, and confers no rights.  "
	$Message_About = $Message_About += "Use of any portion or all of this script are subject to the terms specified at https://www.microsoft.com/en-us/legal/intellectualproperty/copyright/default.aspx."
	$Title_About = "About Exchange Data Collector (ExDC)"
	#$MessageBox_About = [Windows.Forms.MessageBox]::Show($Message_About, $Title_About, [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Information)
	[Windows.Forms.MessageBox]::Show($Message_About, $Title_About, [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Information)
}

#endregion "Main Menu" Events

#region "Step1 - Targets" Events
$handler_btn_Step1_DC_Discover=
{
	Disable-AllTargetsButtons
    $status_Step1.Text = "Step 1 Status: Running"
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "ExDC"
	try{$EventLog.WriteEntry("Starting ExDC Step 1 - Discover Domain Controllers","Information", 10)} catch{}
    & ".\ExDC_Scripts\Core_Build_DC.ps1" -location $location
	if ((Test-Path ".\dc.txt") -eq $true)
	{
		# Empty the listboxes
		$clb_Step1_DC_List.items.clear()
		$array_DC_Filtered = $null
        $File_Location = $location + "\dc.txt"
	    $array_DC = @([System.IO.File]::ReadAllLines($File_Location))
	    foreach ($member_DC in $array_DC)
		{
			if ($member_DC -ne "")
			{
				[array]$array_DC_Filtered += $member_DC
			}
		}
	    $intDCTotal = $array_DC_Filtered.length
		foreach ($member_DC_Filtered in $array_DC_Filtered)
	    {
			$clb_Step1_DC_List.items.add($member_DC_Filtered)
		}
		For ($i=0;$i -le ($intDCTotal - 1);$i++)
		{
			$clb_Step1_DC_List.SetItemChecked($i,$true)
		}
		$txt_DCTotal.Text = "Domain Controller count = " + $intDCTotal
		$txt_DCTotal.visible = $true
	}
	else
	{
		write-host	"The file dc.txt is not present.  Run Discover to create the file."
		$status_Step1.Text = "Step 1 Status: Failed - dc.txt file not found.  Run Discover to create the file."
	}
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "ExDC"
	try{$EventLog.WriteEntry("Ending ExDC Step 1 - Discover Domain Controllers","Information", 11)} Catch{}
	Enable-AllTargetsButtons
}

$handler_btn_Step1_DC_Populate=
{
	Import-TargetsDc
}

$handler_btn_Step1_DC_CheckAll=
{
	Enable-TargetsDc
}

$handler_btn_Step1_DC_UncheckAll=
{
	Disable-TargetsDc
}

$handler_btn_Step1_Ex_Discover=
{
	Disable-AllTargetsButtons
    $status_Step1.Text = "Step 1 Status: Running"
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "ExDC"
	try{$EventLog.WriteEntry("Starting ExDC Step 1 - Discover Exchange servers","Information", 10)} catch{}
    & ".\ExDC_Scripts\Core_Build_Ex.ps1" -location $location
	if ((Test-Path ".\exchange.txt") -eq $true)
	{
		# Empty the listboxes
		$clb_Step1_Ex_List.items.clear()
		$array_Ex_Filtered = $null
        $File_Location = $location + "\exchange.txt"
	    $array_Ex = @([System.IO.File]::ReadAllLines($File_Location))
	    foreach ($member_Ex in $array_Ex)
		{
			if ($member_Ex -ne "")
			{
				[array]$array_Ex_Filtered += $member_Ex
			}
		}
	    $intExTotal = $array_Ex_Filtered.length
		foreach ($member_Ex_Filtered in $array_Ex_Filtered)
	    {
			$clb_Step1_Ex_List.items.add($member_Ex_Filtered)
		}
		For ($i=0;$i -le ($intExTotal - 1);$i++)
		{
			$clb_Step1_Ex_List.SetItemChecked($i,$true)
		}
		$txt_ExchTotal.Text = "Exchange server count = " + $intExTotal
		$txt_ExchTotal.visible = $true
	    $status_Step1.Text = "Step 1 Status: Idle"
	}
	else
	{
		write-host	"The file exchange.txt is not present.  Run Discover to create the file."
		$status_Step1.Text = "Step 1 Status: Failed - exchange.txt file not found.  Run Discover to create the file."
	}
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "ExDC"
	try{$EventLog.WriteEntry("Ending ExDC Step 1 - Discover Exchange servers","Information", 11)} catch{}
	Enable-AllTargetsButtons
}

$handler_btn_Step1_Ex_Populate=
{
	Import-TargetsEx
}

$handler_btn_Step1_Ex_CheckAll=
{
	Enable-TargetsEx
}

$handler_btn_Step1_Ex_UncheckAll=
{
	Disable-TargetsEx
}

$handler_btn_Step1_Nodes_Discover=
{
	Disable-AllTargetsButtons
    $status_Step1.Text = "Step 1 Status: Running"
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "ExDC"
	try{$EventLog.WriteEntry("Starting ExDC Step 1 - Discover Exchange nodes","Information", 10)} catch{}
	if ((Test-Path ".\ClusterNodes.txt") -eq $true)
	{
		write-host "ClusterNodes.txt is already present in this folder." -ForegroundColor Red
		write-host "Loading values from text file that is present." -ForegroundColor Red
		$status_Step1.text = "Step 1 Status: Failed - ClusterNodes.txt is already present. Loading values from existing file."
	}
	else
	{
		$status_Step1.Text = "Step 1 Status: Running - Collecting list of Exchange nodes"
		write-host "Finding cluster nodes in the environment."
		write-host "This could take several minutes."
		$PS_Loc = "$location\ExDC_Scripts\Core_Build_Nodes.ps1"
		Start-Job -ScriptBlock {param($a,$b,$c) Powershell.exe -NoProfile -file $a $b $c} -ArgumentList @($PS_Loc,$location,$session_0) -Name "Core_Build_Nodes.ps1"
		#$strStartJob = Start-Job -ScriptBlock {param($a,$b,$c) Powershell.exe -NoProfile -file $a $b $c} -ArgumentList @($PS_Loc,$location,$session_0) -Name "Core_Build_Nodes.ps1"
#		$strStartJob = Start-Job -FilePath ".\ExDC_Scripts\Core_Build_Nodes.ps1" -argument @($location,$session_0) -Name "Core_Build_Nodes.ps1"
		Update-ExDCJobCount 1 15
	}
	if ((Test-Path ".\ClusterNodes.txt") -eq $true)
	{
		# Empty the listboxes
		$clb_Step1_Nodes_List.items.clear()

		$array_Nodes_Filtered = $null
	    $File_Location = $location + "\ClusterNodes.txt"
        $array_Nodes = @([System.IO.File]::ReadAllLines($File_Location))
	    foreach ($member_Nodes in $array_Nodes)
		{
			if ($member_Nodes -ne "")
			{
				[array]$array_Nodes_Filtered += $member_Nodes
			}
		}
	    $intNodesTotal = $array_Nodes_Filtered.length
	    if ($intNodesTotal -gt 0)
		{
			foreach ($member_Nodes_Filtered in $array_Nodes_Filtered)
		    {
				$clb_Step1_Nodes_List.items.add($member_Nodes_Filtered)
			}
			For ($i=0;$i -le ($intNodesTotal - 1);$i++)
			{
				$clb_Step1_Nodes_List.SetItemChecked($i,$true)
			}
		}
		else
		{
			$intNodesTotal = 0
		}
		$txt_NodesTotal.Text = "Exchange node count = " + $intNodesTotal
		$txt_NodesTotal.visible = $true
	    $status_Step1.Text = "Step 1 Status: Idle"
	}
	else
	{
		write-host	"The file ClusterNodes.txt is not present.  Run Discover to create the file."
		$status_Step1.Text = "Step 1 Status: Failed - ClusterNodes.txt file not found.  Run Discover to create the file."
	}
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "ExDC"
	try{$EventLog.WriteEntry("Ending ExDC Step 1 - Discover Exchange nodes","Information", 11)} catch{}
	Enable-AllTargetsButtons
}

$handler_btn_Step1_Nodes_Populate=
{
	 Import-TargetsNodes
}

$handler_btn_Step1_Nodes_CheckAll=
{
	Enable-TargetsNodes
}

$handler_btn_Step1_Nodes_UncheckAll=
{
	Disable-TargetsNodes
}

$handler_btn_Step1_Mailboxes_Discover=
{
	Disable-AllTargetsButtons
    $status_Step1.Text = "Step 1 Status: Running"
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "ExDC"
	try{$EventLog.WriteEntry("Starting ExDC Step 1 - Discover mailboxes","Information", 10)} catch{}
	$Mailbox_outputfile = ".\Mailbox.txt"
	if ($Exchange2007Powershell -eq $true)
	{
		try
		{([Microsoft.Exchange.Data.Directory.AdminSessionADSettings]::Instance).ViewEntireForest = $true}
		catch{}
	}
	if ($Exchange2010Powershell -eq $true)
	{
	    try
		{Set-AdServerSettings -ViewEntireForest $true}
		catch{}
	}
	if ((Test-Path ".\mailbox.txt") -eq $true)
	{
	    $status_Step1.Text = "Step 1 Status: Failed - mailbox.txt already present.  Please remove and rerun or select Populate."
		write-host "Mailbox.txt is already present in this folder." -ForegroundColor Red
		write-host "Loading values from text file that is present." -ForegroundColor Red
	}
	else
	{
		New-Item $Mailbox_outputfile -type file -Force
	    $MailboxList = @()
			# Old way we got the mailbox list
			#get-mailbox -resultsize unlimited | foreach `
			#{
			#	$MailboxList += $_.alias
			#}
		# Start GC Search method
		# faster than get-mailbox
		$strFilter = $objRootDse = $objRootDomain = $objDomain = $objSearcher = `
		#$colProplist = $colResults = $objResult = $objItem = $null
		$colProplist = $colResults = $objResult = $null
		# RecipientTypeDetails
        # System Attendant Mailbox 8192
        # System Mailbox 16384
        # Arbitration Mailbox 8388608
        # Discovery Mailbox 536870912
        # Health Mailbox 549755813888
        # System Mailbox 4398046511104
		$strFilter = "(&(objectCategory=User)(homeMDB=*)(mailnickname=*)(!msexchrecipienttypedetails=8192)(!msexchrecipienttypedetails=16384)(!msexchrecipienttypedetails=8388608)(!msexchrecipienttypedetails=536870912)(!msexchrecipienttypedetails=549755813888)(!msexchrecipienttypedetails=4398046511104))"
		#$strFilter = "(&(objectCategory=User)(homeMDB=*)(mailnickname=*)(!cn=SystemMailbox{*})(!cn=FederatedEmail.*)(!cn=HealthMailbox*))"
		$objRootDse = New-Object System.DirectoryServices.DirectoryEntry("LDAP://RootDSE")
		$objRootDomain = 'GC://' + $objRootDse.rootDomainNamingContext
		$objDomain = New-Object System.DirectoryServices.DirectoryEntry $objRootDomain
		$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
		$objSearcher.SearchRoot = $objDomain
		$objSearcher.PageSize = 1000
		$objSearcher.Filter = $strFilter
		$objSearcher.SearchScope = "Subtree"
		$colProplist = "mail"
		foreach ($i in $colPropList){$objSearcher.PropertiesToLoad.Add($i)}
		$colResults = $objSearcher.FindAll()
		foreach ($objResult in $colResults)
		{
			$MailboxList += $objResult.Properties.mail
		}
		# End GC Search method
	    $MailboxListSorted = $MailboxList | sort-object
		$MailboxListSorted | out-file $Mailbox_outputfile -append
		$status_Step1.Text = "Step 1 Status: Idle"
	}
    $File_Location = $location + "\mailbox.txt"
	if ((Test-Path $File_Location) -eq $true)
	{
	    $array_Mailboxes = @(([System.IO.File]::ReadAllLines($File_Location)) | sort-object -Unique)
		$intMailboxTotal = 0
		$clb_Step1_Mailboxes_List.items.clear()
	    foreach ($member_Mailbox in $array_Mailboxes | where-object {$_ -ne ""})
	    {
	        $clb_Step1_Mailboxes_List.items.add($member_Mailbox)
			$intMailboxTotal++
	    }
		For ($i=0;$i -le ($intMailboxTotal - 1);$i++)
		{
			$clb_Step1_Mailboxes_List.SetItemChecked($i,$true)
		}
		$txt_MailboxesTotal.Text = "Mailbox count = " + $intMailboxTotal
		$txt_MailboxesTotal.visible = $true
	}
	else
	{
		write-host	"The file mailbox.txt is not present.  Run Discover to create the file."
		$status_Step1.Text = "Step 1 Status: Failed - mailbox.txt file not found.  Run Discover to create the file."
	}
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "ExDC"
	try{$EventLog.WriteEntry("Ending ExDC Step 1 - Discover mailboxes","Information", 11)} catch{}
	Enable-AllTargetsButtons
}

$handler_btn_Step1_Mailboxes_Populate=
{
	 Import-TargetsMailboxes
}

$handler_btn_Step1_Mailboxes_CheckAll=
{
	Enable-TargetsMailbox
}

$handler_btn_Step1_Mailboxes_UncheckAll=
{
	Disable-TargetsMailbox
}

#endregion "Step1 - Targets" Events

#Region "Step2" Events
$handler_rb_Step2_Template_1=
{
	# Uncheck all other radio buttons
	$rb_Step2_Template_2.Checked = $false
	$rb_Step2_Template_3.Checked = $false
	$rb_Step2_Template_4.Checked = $false
	#Load the templates
	#Don't load cluster if tab isn't there
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2003orEarlier -eq $true) -or ($NoEMS -eq $true) -or ($NoGUI -eq $true))
	{
		try{& ".\ExDC_Scripts\Core_Parse_Ini_File.ps1" -IniFile ".\Templates\Template_Recommended_INI_cluster.ini"} catch{}
	}
	if ($NoEMS -eq $false)
	{
		try{& ".\ExDC_Scripts\Core_Parse_Ini_File.ps1" -IniFile ".\Templates\Template_Recommended_INI_ExOrg.ini"} catch{}
	}
	try{& ".\ExDC_Scripts\Core_Parse_Ini_File.ps1" -IniFile ".\Templates\Template_Recommended_INI_server.ini"} catch{}
}

$handler_rb_Step2_Template_2=
{
	# Uncheck all other radio buttons
	$rb_Step2_Template_1.Checked = $false
	$rb_Step2_Template_3.Checked = $false
	$rb_Step2_Template_4.Checked = $false
	#Load the templates
	#Don't load cluster if tab isn't there
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2003orEarlier -eq $true) -or ($NoEMS -eq $true) -or ($NoGUI -eq $true))
	{
		try{& ".\ExDC_Scripts\Core_Parse_Ini_File.ps1" -IniFile ".\Templates\Template_All_INI_cluster.ini"} catch{}
	}
	if ($NoEMS -eq $false)
	{
		try{& ".\ExDC_Scripts\Core_Parse_Ini_File.ps1" -IniFile ".\Templates\Template_All_INI_ExOrg.ini"} catch{}
	}
	try{& ".\ExDC_Scripts\Core_Parse_Ini_File.ps1" -IniFile ".\Templates\Template_All_INI_server.ini"} catch{}
}

$handler_rb_Step2_Template_3=
{
	# Uncheck all other radio buttons
	$rb_Step2_Template_1.Checked = $false
	$rb_Step2_Template_2.Checked = $false
	$rb_Step2_Template_4.Checked = $false
	# Since this is the Environmental Doc template, warn if no EMS
	if ($NoEMS -eq $true)
	{
		write-host "This template is designed to run with the Exchange cmdlets.  NoEMS switch detected." -foregroundcolor yellow
		write-host "Data collection will be incomplete." -foregroundcolor yellow
	}
	#Load the templates
	#Don't load cluster if tab isn't there
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2003orEarlier -eq $true) -or ($NoEMS -eq $true) -or ($NoGUI -eq $true))
	{
		try{& ".\ExDC_Scripts\Core_Parse_Ini_File.ps1" -IniFile ".\Templates\Template_Minimum_INI_cluster.ini"} catch{}
	}
	if ($NoEMS -eq $false)
	{
		try{& ".\ExDC_Scripts\Core_Parse_Ini_File.ps1" -IniFile ".\Templates\Template_Minimum_INI_ExOrg.ini"} catch {}
	}
	& ".\ExDC_Scripts\Core_Parse_Ini_File.ps1" -IniFile ".\Templates\Template_Minimum_INI_server.ini"
}

$handler_rb_Step2_Template_4=
{
	# Uncheck all other radio buttons
	$rb_Step2_Template_1.Checked = $false
	$rb_Step2_Template_2.Checked = $false
	$rb_Step2_Template_3.Checked = $false
	#Load the templates
	#Don't load cluster if tab isn't there
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2003orEarlier -eq $true) -or ($NoEMS -eq $true) -or ($NoGUI -eq $true))
	{
		try{& ".\ExDC_Scripts\Core_Parse_Ini_File.ps1" -IniFile ".\Templates\Template_Custom1_INI_cluster.ini"} catch{}
	}
	if ($NoEMS -eq $false)
	{
		try{& ".\ExDC_Scripts\Core_Parse_Ini_File.ps1" -IniFile ".\Templates\Template_Custom1_INI_ExOrg.ini"} catch{}
	}
	try{& ".\ExDC_Scripts\Core_Parse_Ini_File.ps1" -IniFile ".\Templates\Template_Custom1_INI_server.ini"} catch {}
}

#Endregion "Step2" Events

#region "Step3 - Tests" Events
$handler_btn_Step3_Execute_Click=
{
	Start-Execute
}

$handler_btn_Step3_DC_CheckAll_Click=
{
	Set-AllFunctionsDc -Check $true
}

$handler_btn_Step3_DC_UncheckAll_Click=
{
	Set-AllFunctionsDc -Check $False
}

$handler_btn_Step3_Ex_CheckAll_Click=
{
	Set-AllFunctionsEx -Check $true
}

$handler_btn_Step3_Ex_UncheckAll_Click=
{
	Set-AllFunctionsEx -Check $False
}

$handler_btn_Step3_Cluster_CheckAll_Click=
{
	Set-AllFunctionsCluster -Check $true
}

$handler_btn_Step3_Cluster_UncheckAll_Click=
{
	Set-AllFunctionsCluster -Check $False
}

$handler_btn_Step3_ClientAccess_CheckAll_Click=
{
	Set-AllFunctionsClientAccess -Check $true
}

$handler_btn_Step3_ClientAccess_UncheckAll_Click=
{
	Set-AllFunctionsClientAccess -Check $False
}

$handler_btn_Step3_Global_CheckAll_Click=
{
	Set-AllFunctionsGlobal -Check $true
}

$handler_btn_Step3_Global_UncheckAll_Click=
{
	Set-AllFunctionsGlobal -Check $false
}

$handler_btn_Step3_Recipient_CheckAll_Click=
{
	Set-AllFunctionsRecipient -Check $true
}

$handler_btn_Step3_Recipient_UncheckAll_Click=
{
	Set-AllFunctionsRecipient -Check $False
}

$handler_btn_Step3_Transport_CheckAll_Click=
{
	Set-AllFunctionsTransport -Check $true
}

$handler_btn_Step3_Transport_UncheckAll_Click=
{
	Set-AllFunctionsTransport -Check $False
}

$handler_btn_Step3_Um_CheckAll_Click=
{
	Set-AllFunctionsUm -Check $true
}

$handler_btn_Step3_Um_UncheckAll_Click=
{
	Set-AllFunctionsUm -Check $False
}

$handler_btn_Step3_Misc_CheckAll_Click=
{
	Set-AllFunctionsMisc -Check $true
}

$handler_btn_Step3_Misc_UncheckAll_Click=
{
	Set-AllFunctionsMisc -Check $False
}
#endregion "Step3 - Tests" Events

#region "Step4 - Reporting" Events
$handler_btn_Step4_Assemble_Click=
{
	$btn_Step4_Assemble.enabled = $false
    $status_Step4.Text = "Step 4 Status: Running"
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "ExDC"
	try{$EventLog.WriteEntry("Starting ExDC Step 4","Information", 40)} catch{}
	#Minimize form to the back to expose the Powershell window when starting Step 4
	$form1.WindowState = "minimized"
	write-host "ExDC Form minimized." -ForegroundColor Green
	if ((Test-Path registry::HKey_Classes_Root\Excel.Application\CurVer) -eq $true)
	{
		if ($chk_Step4_DC_Report.checked -eq $true)
		{
			write-host "-- Starting to assemble the DC Spreadsheet"
				.\ExDC_Scripts\Core_assemble_dc_Excel.ps1 -RunLocation $location
				write-host "---- Completed the DC Spreadsheet" -ForegroundColor Green
		}
		if ($chk_Step4_Ex_Report.checked -eq $true)
		{
			write-host "-- Starting to assemble the Exchange Server Spreadsheet"
				.\ExDC_Scripts\Core_assemble_exch_Excel.ps1 -RunLocation $location
				write-host "---- Completed the Exchange Spreadsheet" -ForegroundColor Green
		}
		if ($chk_Step4_ExOrg_Report.checked -eq $true)
		{
			write-host "-- Starting to assemble the Exchange Organization Spreadsheet"
				.\ExDC_Scripts\Core_assemble_exorg_Excel.ps1 -RunLocation $location
				write-host "---- Completed the Exchange Organization Spreadsheet" -ForegroundColor Green
		}
	}
	else
	{
		write-host "Excel does not appear to be installed on this server."
		$EventLog = New-Object System.Diagnostics.EventLog('Application')
		$EventLog.MachineName = "."
		$EventLog.Source = "ExDC"
		try{$EventLog.WriteEntry("Excel does not appear to be installed on this server.","Warning", 49)} catch{}
	}
	if ((Test-Path registry::HKey_Classes_Root\Word.Application\CurVer) -eq $true)
	{
		if ($chk_Step4_Exchange_Environment_Doc.checked -eq $true)
		{
			write-host "-- Starting to assemble the Exchange Documentation using Word"
				.\ExDC_Scripts\Core_Assemble_ExDoc_Word.ps1 -RunLocation $location
				write-host "---- Completed the Exchange Documentation using Word" -ForegroundColor Green
		}
	}
	else
	{
		write-host "Word does not appear to be installed on this server."
		$EventLog = New-Object System.Diagnostics.EventLog('Application')
		$EventLog.MachineName = "."
		$EventLog.Source = "ExDC"
		try{$EventLog.WriteEntry("Word does not appear to be installed on this server.","Warning", 49)} catch{}
	}
	write-host "Restoring ExDC Form to normal." -ForegroundColor Green
	$form1.WindowState = "normal"
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "ExDC"
	try{$EventLog.WriteEntry("Ending ExDC Step 4","Information", 41)} catch{}
	$status_Step4.Text = "Step 4 Status: Idle"
    $btn_Step4_Assemble.enabled = $true
}
#endregion "Step4 - Reporting" Events

#region *** Events ***

#endregion *** Events ***

$OnLoadForm_StateCorrection=
{$form1.WindowState = $InitialFormWindowState}

#region *** Build Form ***

#Region Form Main
# Reusable fonts
	$font_Calibri_8pt_normal = 	New-Object System.Drawing.Font("Calibri",7.8,0,3,0)
	$font_Calibri_10pt_normal = New-Object System.Drawing.Font("Calibri",9.75,0,3,1)
	$font_Calibri_12pt_normal = New-Object System.Drawing.Font("Calibri",12,0,3,1)
	$font_Calibri_14pt_normal = New-Object System.Drawing.Font("Calibri",14.25,0,3,1)
	$font_Calibri_10pt_bold = 	New-Object System.Drawing.Font("Calibri",9.75,1,3,1)
# Reusable padding
	$System_Windows_Forms_Padding_Reusable = New-Object System.Windows.Forms.Padding
	$System_Windows_Forms_Padding_Reusable.All = 3
	$System_Windows_Forms_Padding_Reusable.Bottom = 3
	$System_Windows_Forms_Padding_Reusable.Left = 3
	$System_Windows_Forms_Padding_Reusable.Right = 3
	$System_Windows_Forms_Padding_Reusable.Top = 3
# Reusable button
	$System_Drawing_Size_buttons = New-Object System.Drawing.Size
	$System_Drawing_Size_buttons.Height = 38
	$System_Drawing_Size_buttons.Width = 110
# Reusable status
	$System_Drawing_Size_Status = New-Object System.Drawing.Size
	$System_Drawing_Size_Status.Height = 22
	$System_Drawing_Size_Status.Width = 651
	$System_Drawing_Point_Status = New-Object System.Drawing.Point
	$System_Drawing_Point_Status.X = 3
	$System_Drawing_Point_Status.Y = 653
# Reusable tabs
	$System_Drawing_Size_tab_1 = New-Object System.Drawing.Size
	$System_Drawing_Size_tab_1.Height = 678
	$System_Drawing_Size_tab_1.Width = 700 #657
	$System_Drawing_Size_tab_2 = New-Object System.Drawing.Size
	$System_Drawing_Size_tab_2.Height = 678
	$System_Drawing_Size_tab_2.Width = 1000
# Reusable checkboxes
	$System_Drawing_Size_Reusable_chk = New-Object System.Drawing.Size
	$System_Drawing_Size_Reusable_chk.Height = 20
	$System_Drawing_Size_Reusable_chk.Width = 225
	$System_Drawing_Size_Reusable_chk_long = New-Object System.Drawing.Size
	$System_Drawing_Size_Reusable_chk_long.Height = 20
	$System_Drawing_Size_Reusable_chk_long.Width = 400

# Main Form
$form1.BackColor = [System.Drawing.Color]::FromArgb(255,169,169,169)
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 718
	$System_Drawing_Size.Width = 665
$form1.ClientSize = $System_Drawing_Size
$form1.MaximumSize = $System_Drawing_Size
$form1.Font = $font_Calibri_10pt_normal
$form1.FormBorderStyle = 2
$form1.MaximizeBox = $False
$form1.Name = "form1"
$form1.ShowIcon = $False
$form1.StartPosition = 1
$form1.Text = "Exchange Data Collector v4.0.2"

# Main Tabs
$tab_Master.Appearance = 2
$tab_Master.Dock = 5
$tab_Master.Font = $font_Calibri_14pt_normal
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 32
	$System_Drawing_Size.Width = 100
$tab_Master.ItemSize = $System_Drawing_Size
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 0
	$System_Drawing_Point.Y = 0
$tab_Master.Location = $System_Drawing_Point
$tab_Master.Name = "tab_Master"
$tab_Master.SelectedIndex = 0
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 718
	$System_Drawing_Size.Width = 665
$tab_Master.Size = $System_Drawing_Size
$tab_Master.SizeMode = "filltoright"
$tab_Master.TabIndex = 12
$form1.Controls.Add($tab_Master)

# Menu Strip
$Menu_Main.Location = new-object System.Drawing.Point(0, 0)
$Menu_Main.Name = "MainMenu"
$Menu_Main.Size = new-object System.Drawing.Size(1151, 24)
$Menu_Main.TabIndex = 0
$Menu_Main.Text = "Main Menu"
$form1.Controls.add($Menu_Main)
[Void]$Menu_File.DropDownItems.Add($Submenu_LoadTargets)
[Void]$Menu_File.DropDownItems.Add($Submenu_PackageLogs)
[Void]$Menu_Main.items.add($Menu_File)
[Void]$Menu_Toggle.DropDownItems.Add($Submenu_Targets_CheckAll)
[Void]$Menu_Toggle.DropDownItems.Add($Submenu_Targets_UnCheckAll)
[Void]$Menu_Toggle.DropDownItems.Add($Submenu_Tests_CheckAll)
[Void]$Menu_Toggle.DropDownItems.Add($Submenu_Tests_UnCheckAll)
[Void]$Menu_Main.items.add($Menu_Toggle)
[Void]$Menu_Help.DropDownItems.Add($Submenu_Help)
[Void]$Menu_Help.DropDownItems.Add($Submenu_About)
[Void]$Menu_Main.items.add($Menu_Help)
$Submenu_LoadTargets.add_click($handler_Submenu_LoadTargets)
$Submenu_PackageLogs.add_click($handler_Submenu_PackageLogs)
$Submenu_Targets_CheckAll.add_click($handler_Submenu_Targets_CheckAll)
$Submenu_Targets_UnCheckAll.add_click($handler_Submenu_Targets_UnCheckAll)
$Submenu_Tests_CheckAll.add_click($handler_Submenu_Tests_CheckAll)
$Submenu_Tests_UnCheckAll.add_click($handler_Submenu_Tests_UnCheckAll)
$Submenu_Help.add_click($handler_Submenu_Help)
$Submenu_About.add_click($handler_Submenu_About)
#EndRegion Form Main

#Region "Step1 - Targets"

#Region Step1 Main
# Reusable text box in Step1
	$System_Drawing_Size_Step1_text_box = New-Object System.Drawing.Size
	$System_Drawing_Size_Step1_text_box.Height = 27
	$System_Drawing_Size_Step1_text_box.Width = 400
# Reusable label in Step1
	$System_Drawing_Size_Step1_label = New-Object System.Drawing.Size
	$System_Drawing_Size_Step1_label.Height = 20
	$System_Drawing_Size_Step1_label.Width = 200
# Reusable Listbox in Step1
	$System_Drawing_Size_Step1_Listbox = New-Object System.Drawing.Size
	$System_Drawing_Size_Step1_Listbox.Height = 384
	$System_Drawing_Size_Step1_Listbox.Width = 200
# Reusable boxes in Step1 Tabs
	$System_Drawing_Size_Step1_box = New-Object System.Drawing.Size
	$System_Drawing_Size_Step1_box.Height = 482
	$System_Drawing_Size_Step1_box.Width = 536
# Reusable check buttons in Step1 tabs
	$System_Drawing_Size_Step1_btn = New-Object System.Drawing.Size
	$System_Drawing_Size_Step1_btn.Height = 25
	$System_Drawing_Size_Step1_btn.Width = 150
# Reusable check list boxes in Step1 tabs
	$System_Drawing_Size_Step1_clb = New-Object System.Drawing.Size
	$System_Drawing_Size_Step1_clb.Height = 350
	$System_Drawing_Size_Step1_clb.Width = 400
	$System_Drawing_Point_Step1_clb = New-Object System.Drawing.Point
	$System_Drawing_Point_Step1_clb.X = 50
	$System_Drawing_Point_Step1_clb.Y = 50
# Reusable Discover/populate buttons in Step1 tabs
	$System_Drawing_Point_Step1_Discover = New-Object System.Drawing.Point
	$System_Drawing_Point_Step1_Discover.X = 50
	$System_Drawing_Point_Step1_Discover.Y = 15
	$System_Drawing_Point_Step1_Populate = New-Object System.Drawing.Point
	$System_Drawing_Point_Step1_Populate.X = 300
	$System_Drawing_Point_Step1_Populate.Y = 15
# Reusable check/uncheck buttons in Step1 tabs
	$System_Drawing_Point_Step1_CheckAll = New-Object System.Drawing.Point
	$System_Drawing_Point_Step1_CheckAll.X = 50
	$System_Drawing_Point_Step1_CheckAll.Y = 450
	$System_Drawing_Point_Step1_UncheckAll = New-Object System.Drawing.Point
	$System_Drawing_Point_Step1_UncheckAll.X = 300
	$System_Drawing_Point_Step1_UncheckAll.Y = 450
$tab_Step1.BackColor = [System.Drawing.Color]::FromArgb(0,255,255,255)
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 36
$tab_Step1.Location = $System_Drawing_Point
$tab_Step1.Name = "tab_Step1"
$tab_Step1.Padding = $System_Windows_Forms_Padding_Reusable
$tab_Step1.TabIndex = 0
$tab_Step1.Text = "  Targets  "
$tab_Step1.Size = $System_Drawing_Size_tab_1
$tab_Master.Controls.Add($tab_Step1)
$btn_Step1_Discover.Font = $font_Calibri_14pt_normal
$btn_Step1_Discover.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 20
	$System_Drawing_Point.Y = 15
$btn_Step1_Discover.Location = $System_Drawing_Point
$btn_Step1_Discover.Name = "btn_Step1_Discover"
$btn_Step1_Discover.Size = $System_Drawing_Size_buttons
$btn_Step1_Discover.TabIndex = 0
$btn_Step1_Discover.Text = "Discover"
$btn_Step1_Discover.Visible = $false
$btn_Step1_Discover.UseVisualStyleBackColor = $True
$btn_Step1_Discover.add_Click($handler_btn_Step1_Discover_Click)
$tab_Step1.Controls.Add($btn_Step1_Discover)
$btn_Step1_Populate.Font = $font_Calibri_14pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 200
	$System_Drawing_Point.Y = 15
$btn_Step1_Populate.Location = $System_Drawing_Point
$btn_Step1_Populate.Name = "btn_Step1_Populate"
$btn_Step1_Populate.Size = $System_Drawing_Size_buttons
$btn_Step1_Populate.TabIndex = 9
$btn_Step1_Populate.Text = "Load from File"
$btn_Step1_Populate.Visible = $false
$btn_Step1_Populate.UseVisualStyleBackColor = $True
$btn_Step1_Populate.add_Click($handler_btn_Step1_Populate_Click)
$tab_Step1.Controls.Add($btn_Step1_Populate)
$tab_Step1_Master.Font = $font_Calibri_12pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 20
	$System_Drawing_Point.Y = 60
$tab_Step1_Master.Location = $System_Drawing_Point
$tab_Step1_Master.Name = "tab_Step1_Master"
$tab_Step1_Master.SelectedIndex = 0
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 525
	$System_Drawing_Size.Width = 550
$tab_Step1_Master.Size = $System_Drawing_Size
$tab_Step1_Master.TabIndex = 11
$tab_Step1.Controls.Add($tab_Step1_Master)
$status_Step1.Font = $font_Calibri_10pt_normal
$status_Step1.Location = $System_Drawing_Point_Status
$status_Step1.Name = "status_Step1"
$status_Step1.Size = $System_Drawing_Size_Status
$status_Step1.TabIndex = 2
$status_Step1.Text = "Step 1 Status"
$tab_Step1.Controls.Add($status_Step1)
#EndRegion Step1 Main

#Region Step1 Domain Controller tab
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 33
$tab_Step1_DC.Location = $System_Drawing_Point
$tab_Step1_DC.Name = "tab_Step1_DC"
$tab_Step1_DC.Padding = $System_Windows_Forms_Padding_Reusable
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 488
	$System_Drawing_Size.Width = 542
$tab_Step1_DC.Size = $System_Drawing_Size
$tab_Step1_DC.TabIndex = 0
$tab_Step1_DC.Text = "Domain Controllers"
$tab_Step1_DC.UseVisualStyleBackColor = $True
$tab_Step1_Master.Controls.Add($tab_Step1_DC)
$bx_DC_List.Dock = 5
$bx_DC_List.Font = $font_Calibri_10pt_bold
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 3
	$System_Drawing_Point.Y = 3
$bx_DC_List.Location = $System_Drawing_Point
$bx_DC_List.Name = "bx_DC_List"
$bx_DC_List.Size = $System_Drawing_Size_Step1_box
$bx_DC_List.TabIndex = 7
$bx_DC_List.TabStop = $False
$tab_Step1_DC.Controls.Add($bx_DC_List)
$btn_Step1_DC_Discover.Font = $font_Calibri_10pt_normal
$btn_Step1_DC_Discover.Location = $System_Drawing_Point_Step1_Discover
$btn_Step1_DC_Discover.Name = "btn_Step1_DC_Discover"
$btn_Step1_DC_Discover.Size = $System_Drawing_Size_Step1_btn
$btn_Step1_DC_Discover.TabIndex = 9
$btn_Step1_DC_Discover.Text = "Discover"
$btn_Step1_DC_Discover.UseVisualStyleBackColor = $True
$btn_Step1_DC_Discover.add_Click($handler_btn_Step1_DC_Discover)
$bx_DC_List.Controls.Add($btn_Step1_DC_Discover)
$btn_Step1_DC_Populate.Font = $font_Calibri_10pt_normal
$btn_Step1_DC_Populate.Location = $System_Drawing_Point_Step1_Populate
$btn_Step1_DC_Populate.Name = "btn_Step1_DC_Populate"
$btn_Step1_DC_Populate.Size = $System_Drawing_Size_Step1_btn
$btn_Step1_DC_Populate.TabIndex = 10
$btn_Step1_DC_Populate.Text = "Load from File"
$btn_Step1_DC_Populate.UseVisualStyleBackColor = $True
$btn_Step1_DC_Populate.add_Click($handler_btn_Step1_DC_Populate)
$bx_DC_List.Controls.Add($btn_Step1_DC_Populate)
$clb_Step1_DC_List.Font = $font_Calibri_10pt_normal
$clb_Step1_DC_List.Location = $System_Drawing_Point_Step1_clb
$clb_Step1_DC_List.Name = "clb_Step1_DC_List"
$clb_Step1_DC_List.Size = $System_Drawing_Size_Step1_clb
$clb_Step1_DC_List.TabIndex = 10
$clb_Step1_DC_List.horizontalscrollbar = $true
$clb_Step1_DC_List.CheckOnClick = $true
$bx_DC_List.Controls.Add($clb_Step1_DC_List)
$txt_DCTotal.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 50
	$System_Drawing_Point.Y = 410
$txt_DCTotal.Location = $System_Drawing_Point
$txt_DCTotal.Name = "txt_DCTotal"
$txt_DCTotal.Size = $System_Drawing_Size_Step1_text_box
$txt_DCTotal.TabIndex = 11
$txt_DCTotal.Visible = $False
$bx_DC_List.Controls.Add($txt_DCTotal)
$btn_Step1_DC_CheckAll.Font = $font_Calibri_10pt_normal
$btn_Step1_DC_CheckAll.Location = $System_Drawing_Point_Step1_CheckAll
$btn_Step1_DC_CheckAll.Name = "btn_Step1_DC_CheckAll"
$btn_Step1_DC_CheckAll.Size = $System_Drawing_Size_Step1_btn
$btn_Step1_DC_CheckAll.TabIndex = 9
$btn_Step1_DC_CheckAll.Text = "Check all on this tab"
$btn_Step1_DC_CheckAll.UseVisualStyleBackColor = $True
$btn_Step1_DC_CheckAll.add_Click($handler_btn_Step1_DC_CheckAll)
$bx_DC_List.Controls.Add($btn_Step1_DC_CheckAll)
$btn_Step1_DC_UncheckAll.Font = $font_Calibri_10pt_normal
$btn_Step1_DC_UncheckAll.Location = $System_Drawing_Point_Step1_UncheckAll
$btn_Step1_DC_UncheckAll.Name = "btn_Step1_DC_UncheckAll"
$btn_Step1_DC_UncheckAll.Size = $System_Drawing_Size_Step1_btn
$btn_Step1_DC_UncheckAll.TabIndex = 10
$btn_Step1_DC_UncheckAll.Text = "Uncheck all on this tab"
$btn_Step1_DC_UncheckAll.UseVisualStyleBackColor = $True
$btn_Step1_DC_UncheckAll.add_Click($handler_btn_Step1_DC_UncheckAll)
$bx_DC_List.Controls.Add($btn_Step1_DC_UncheckAll)
#EndRegion Step1 Domain Controller tab

#Region Step1 Exchange server tab
$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 33
$tab_Step1_Ex.Location = $System_Drawing_Point
$tab_Step1_Ex.Name = "tab_Step1_Ex"
$tab_Step1_Ex.Padding = $System_Windows_Forms_Padding_Reusable
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 488
	$System_Drawing_Size.Width = 542
$tab_Step1_Ex.Size = $System_Drawing_Size
$tab_Step1_Ex.TabIndex = 1
$tab_Step1_Ex.Text = "Exchange Servers"
$tab_Step1_Ex.UseVisualStyleBackColor = $True
$tab_Step1_Master.Controls.Add($tab_Step1_Ex)
$bx_Ex_List.Dock = 5
$bx_Ex_List.Font = $font_Calibri_10pt_bold
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 3
	$System_Drawing_Point.Y = 3
$bx_Ex_List.Location = $System_Drawing_Point
$bx_Ex_List.Name = "bx_Ex_List"
$bx_Ex_List.Size = $System_Drawing_Size_Step1_box
$bx_Ex_List.TabIndex = 7
$bx_Ex_List.TabStop = $False
$tab_Step1_Ex.Controls.Add($bx_Ex_List)
$btn_Step1_Ex_Discover.Font = $font_Calibri_10pt_normal
$btn_Step1_Ex_Discover.Location = $System_Drawing_Point_Step1_Discover
$btn_Step1_Ex_Discover.Name = "btn_Step1_Ex_Discover"
$btn_Step1_Ex_Discover.Size = $System_Drawing_Size_Step1_btn
$btn_Step1_Ex_Discover.TabIndex = 9
$btn_Step1_Ex_Discover.Text = "Discover"
$btn_Step1_Ex_Discover.UseVisualStyleBackColor = $True
$btn_Step1_Ex_Discover.add_Click($handler_btn_Step1_Ex_Discover)
$bx_Ex_List.Controls.Add($btn_Step1_Ex_Discover)
$btn_Step1_Ex_Populate.Font = $font_Calibri_10pt_normal
$btn_Step1_Ex_Populate.Location = $System_Drawing_Point_Step1_Populate
$btn_Step1_Ex_Populate.Name = "btn_Step1_Ex_Populate"
$btn_Step1_Ex_Populate.Size = $System_Drawing_Size_Step1_btn
$btn_Step1_Ex_Populate.TabIndex = 10
$btn_Step1_Ex_Populate.Text = "Load from File"
$btn_Step1_Ex_Populate.UseVisualStyleBackColor = $True
$btn_Step1_Ex_Populate.add_Click($handler_btn_Step1_Ex_Populate)
$bx_Ex_List.Controls.Add($btn_Step1_Ex_Populate)
$clb_Step1_Ex_List.Font = $font_Calibri_10pt_normal
$clb_Step1_Ex_List.Location = $System_Drawing_Point_Step1_clb
$clb_Step1_Ex_List.Name = "clb_Step1_Ex_List"
$clb_Step1_Ex_List.Size = $System_Drawing_Size_Step1_clb
$clb_Step1_Ex_List.TabIndex = 10
$clb_Step1_Ex_List.horizontalscrollbar = $true
$clb_Step1_Ex_List.CheckOnClick = $true
$bx_Ex_List.Controls.Add($clb_Step1_Ex_List)
$txt_ExchTotal.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 50
	$System_Drawing_Point.Y = 410
$txt_ExchTotal.Location = $System_Drawing_Point
$txt_ExchTotal.Name = "txt_ExTotal"
$txt_ExchTotal.Size = $System_Drawing_Size_Step1_text_box
$txt_ExchTotal.TabIndex = 11
$txt_ExchTotal.Visible = $False
$bx_Ex_List.Controls.Add($txt_ExchTotal)
$btn_Step1_Ex_CheckAll.Font = $font_Calibri_10pt_normal
$btn_Step1_Ex_CheckAll.Location = $System_Drawing_Point_Step1_CheckAll
$btn_Step1_Ex_CheckAll.Name = "btn_Step1_Ex_CheckAll"
$btn_Step1_Ex_CheckAll.Size = $System_Drawing_Size_Step1_btn
$btn_Step1_Ex_CheckAll.TabIndex = 9
$btn_Step1_Ex_CheckAll.Text = "Check all on this tab"
$btn_Step1_Ex_CheckAll.UseVisualStyleBackColor = $True
$btn_Step1_Ex_CheckAll.add_Click($handler_btn_Step1_Ex_CheckAll)
$bx_Ex_List.Controls.Add($btn_Step1_Ex_CheckAll)
$btn_Step1_Ex_UncheckAll.Font = $font_Calibri_10pt_normal
$btn_Step1_Ex_UncheckAll.Location = $System_Drawing_Point_Step1_UncheckAll
$btn_Step1_Ex_UncheckAll.Name = "btn_Step1_Ex_UncheckAll"
$btn_Step1_Ex_UncheckAll.Size = $System_Drawing_Size_Step1_btn
$btn_Step1_Ex_UncheckAll.TabIndex = 10
$btn_Step1_Ex_UncheckAll.Text = "Uncheck all on this tab"
$btn_Step1_Ex_UncheckAll.UseVisualStyleBackColor = $True
$btn_Step1_Ex_UncheckAll.add_Click($handler_btn_Step1_Ex_UncheckAll)
$bx_Ex_List.Controls.Add($btn_Step1_Ex_UncheckAll)
#EndRegion Step1 Exchange server  tab

#Region Step1 Nodes tab
if (($Exchange2007Powershell -eq $true) -or ($Exchange2003orEarlier -eq $true) -or ($NoEMS -eq $true) -or ($NoGUI -eq $true))
{
	$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = 4
		$System_Drawing_Point.Y = 33
	$tab_Step1_Nodes.Location = $System_Drawing_Point
	$tab_Step1_Nodes.Name = "tab_Step1_Nodes"
	$tab_Step1_Nodes.Padding = $System_Windows_Forms_Padding_Reusable
		$System_Drawing_Size = New-Object System.Drawing.Size
		$System_Drawing_Size.Height = 488
		$System_Drawing_Size.Width = 542
	$tab_Step1_Nodes.Size = $System_Drawing_Size
	$tab_Step1_Nodes.TabIndex = 1
	$tab_Step1_Nodes.Text = "Exchange Nodes"
	$tab_Step1_Nodes.UseVisualStyleBackColor = $True
	$tab_Step1_Master.Controls.Add($tab_Step1_Nodes)
	$btn_Step1_Nodes_Discover.Font = $font_Calibri_10pt_normal
	$btn_Step1_Nodes_Discover.Location = $System_Drawing_Point_Step1_Discover
	$btn_Step1_Nodes_Discover.Name = "btn_Step1_Nodes_Discover"
	$btn_Step1_Nodes_Discover.Size = $System_Drawing_Size_Step1_btn
	$btn_Step1_Nodes_Discover.TabIndex = 9
	$btn_Step1_Nodes_Discover.Text = "Discover"
	$btn_Step1_Nodes_Discover.UseVisualStyleBackColor = $True
	$btn_Step1_Nodes_Discover.add_Click($handler_btn_Step1_Nodes_Discover)
	$bx_Nodes_List.Controls.Add($btn_Step1_Nodes_Discover)
	$btn_Step1_Nodes_Populate.Font = $font_Calibri_10pt_normal
	$btn_Step1_Nodes_Populate.Location = $System_Drawing_Point_Step1_Populate
	$btn_Step1_Nodes_Populate.Name = "btn_Step1_Nodes_Populate"
	$btn_Step1_Nodes_Populate.Size = $System_Drawing_Size_Step1_btn
	$btn_Step1_Nodes_Populate.TabIndex = 10
	$btn_Step1_Nodes_Populate.Text = "Load from File"
	$btn_Step1_Nodes_Populate.UseVisualStyleBackColor = $True
	$btn_Step1_Nodes_Populate.add_Click($handler_btn_Step1_Nodes_Populate)
	$bx_Nodes_List.Controls.Add($btn_Step1_Nodes_Populate)
	$bx_Nodes_List.Dock = 5
	$bx_Nodes_List.Font = $font_Calibri_10pt_bold
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = 3
		$System_Drawing_Point.Y = 3
	$bx_Nodes_List.Location = $System_Drawing_Point
	$bx_Nodes_List.Name = "bx_Nodes_List"
	$bx_Nodes_List.Size = $System_Drawing_Size_Step1_box
	$bx_Nodes_List.TabIndex = 7
	$bx_Nodes_List.TabStop = $False
	$tab_Step1_Nodes.Controls.Add($bx_Nodes_List)
	$clb_Step1_Nodes_List.Font = $font_Calibri_10pt_normal
	$clb_Step1_Nodes_List.Location = $System_Drawing_Point_Step1_clb
	$clb_Step1_Nodes_List.Name = "clb_Step1_Nodes_List"
	$clb_Step1_Nodes_List.Size = $System_Drawing_Size_Step1_clb
	$clb_Step1_Nodes_List.TabIndex = 10
	$clb_Step1_Nodes_List.horizontalscrollbar = $true
	$clb_Step1_Nodes_List.CheckOnClick = $true
	$bx_Nodes_List.Controls.Add($clb_Step1_Nodes_List)
	$txt_NodesTotal.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = 50
		$System_Drawing_Point.Y = 410
	$txt_NodesTotal.Location = $System_Drawing_Point
	$txt_NodesTotal.Name = "txt_NodesTotal"
	$txt_NodesTotal.Size = $System_Drawing_Size_Step1_text_box
	$txt_NodesTotal.TabIndex = 11
	$txt_NodesTotal.Visible = $False
	$bx_Nodes_List.Controls.Add($txt_NodesTotal)
	$btn_Step1_Nodes_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step1_Nodes_CheckAll.Location = $System_Drawing_Point_Step1_CheckAll
	$btn_Step1_Nodes_CheckAll.Name = "btn_Step1_Nodes_CheckAll"
	$btn_Step1_Nodes_CheckAll.Size = $System_Drawing_Size_Step1_btn
	$btn_Step1_Nodes_CheckAll.TabIndex = 9
	$btn_Step1_Nodes_CheckAll.Text = "Check all on this tab"
	$btn_Step1_Nodes_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step1_Nodes_CheckAll.add_Click($handler_btn_Step1_Nodes_CheckAll)
	$bx_Nodes_List.Controls.Add($btn_Step1_Nodes_CheckAll)
	$btn_Step1_Nodes_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step1_Nodes_UncheckAll.Location = $System_Drawing_Point_Step1_UncheckAll
	$btn_Step1_Nodes_UncheckAll.Name = "btn_Step1_Nodes_UncheckAll"
	$btn_Step1_Nodes_UncheckAll.Size = $System_Drawing_Size_Step1_btn
	$btn_Step1_Nodes_UncheckAll.TabIndex = 10
	$btn_Step1_Nodes_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step1_Nodes_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step1_Nodes_UncheckAll.add_Click($handler_btn_Step1_Nodes_UncheckAll)
	$bx_Nodes_List.Controls.Add($btn_Step1_Nodes_UncheckAll)
}
#EndRegion Step1 Nodes tab

#Region Step1 Mailboxes tab
if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true) -or ($NoGUI -eq $true))
{
	$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = 4
		$System_Drawing_Point.Y = 33
	$tab_Step1_Mailboxes.Location = $System_Drawing_Point
	$tab_Step1_Mailboxes.Name = "tab_Step1_Mailboxes"
	$tab_Step1_Mailboxes.Padding = $System_Windows_Forms_Padding_Reusable
		$System_Drawing_Size = New-Object System.Drawing.Size
		$System_Drawing_Size.Height = 488
		$System_Drawing_Size.Width = 542
	$tab_Step1_Mailboxes.Size = $System_Drawing_Size
	$tab_Step1_Mailboxes.TabIndex = 1
	$tab_Step1_Mailboxes.Text = "Mailboxes"
	$tab_Step1_Mailboxes.UseVisualStyleBackColor = $True
	$tab_Step1_Master.Controls.Add($tab_Step1_Mailboxes)
	$btn_Step1_Mailboxes_Discover.Font = $font_Calibri_10pt_normal
	$btn_Step1_Mailboxes_Discover.Location = $System_Drawing_Point_Step1_Discover
	$btn_Step1_Mailboxes_Discover.Name = "btn_Step1_Mailboxes_Discover"
	$btn_Step1_Mailboxes_Discover.Size = $System_Drawing_Size_Step1_btn
	$btn_Step1_Mailboxes_Discover.TabIndex = 9
	$btn_Step1_Mailboxes_Discover.Text = "Discover"
	$btn_Step1_Mailboxes_Discover.UseVisualStyleBackColor = $True
	$btn_Step1_Mailboxes_Discover.add_Click($handler_btn_Step1_Mailboxes_Discover)
	$bx_Mailboxes_List.Controls.Add($btn_Step1_Mailboxes_Discover)
	$btn_Step1_Mailboxes_Populate.Font = $font_Calibri_10pt_normal
	$btn_Step1_Mailboxes_Populate.Location = $System_Drawing_Point_Step1_Populate
	$btn_Step1_Mailboxes_Populate.Name = "btn_Step1_Mailboxes_Populate"
	$btn_Step1_Mailboxes_Populate.Size = $System_Drawing_Size_Step1_btn
	$btn_Step1_Mailboxes_Populate.TabIndex = 10
	$btn_Step1_Mailboxes_Populate.Text = "Load from File"
	$btn_Step1_Mailboxes_Populate.UseVisualStyleBackColor = $True
	$btn_Step1_Mailboxes_Populate.add_Click($handler_btn_Step1_Mailboxes_Populate)
	$bx_Mailboxes_List.Controls.Add($btn_Step1_Mailboxes_Populate)
	$bx_Mailboxes_List.Dock = 5
	$bx_Mailboxes_List.Font = $font_Calibri_10pt_bold
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = 3
		$System_Drawing_Point.Y = 3
	$bx_Mailboxes_List.Location = $System_Drawing_Point
	$bx_Mailboxes_List.Name = "bx_Mailboxes_List"
	$bx_Mailboxes_List.Size = $System_Drawing_Size_Step1_box
	$bx_Mailboxes_List.TabIndex = 7
	$bx_Mailboxes_List.TabStop = $False
	$tab_Step1_Mailboxes.Controls.Add($bx_Mailboxes_List)
	$clb_Step1_Mailboxes_List.Font = $font_Calibri_10pt_normal
	$clb_Step1_Mailboxes_List.Location = $System_Drawing_Point_Step1_clb
	$clb_Step1_Mailboxes_List.Name = "clb_Step1_Mailboxes_List"
	$clb_Step1_Mailboxes_List.Size = $System_Drawing_Size_Step1_clb
	$clb_Step1_Mailboxes_List.TabIndex = 10
	$clb_Step1_Mailboxes_List.horizontalscrollbar = $true
	$clb_Step1_Mailboxes_List.CheckOnClick = $true
	$bx_Mailboxes_List.Controls.Add($clb_Step1_Mailboxes_List)
	$txt_MailboxesTotal.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = 50
		$System_Drawing_Point.Y = 410
	$txt_MailboxesTotal.Location = $System_Drawing_Point
	$txt_MailboxesTotal.Name = "txt_MailboxesTotal"
	$txt_MailboxesTotal.Size = $System_Drawing_Size_Step1_text_box
	$txt_MailboxesTotal.TabIndex = 11
	$txt_MailboxesTotal.Visible = $False
	$bx_Mailboxes_List.Controls.Add($txt_MailboxesTotal)
	$btn_Step1_Mailboxes_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step1_Mailboxes_CheckAll.Location = $System_Drawing_Point_Step1_CheckAll
	$btn_Step1_Mailboxes_CheckAll.Name = "btn_Step1_Mailboxes_CheckAll"
	$btn_Step1_Mailboxes_CheckAll.Size = $System_Drawing_Size_Step1_btn
	$btn_Step1_Mailboxes_CheckAll.TabIndex = 9
	$btn_Step1_Mailboxes_CheckAll.Text = "Check all on this tab"
	$btn_Step1_Mailboxes_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step1_Mailboxes_CheckAll.add_Click($handler_btn_Step1_Mailboxes_CheckAll)
	$bx_Mailboxes_List.Controls.Add($btn_Step1_Mailboxes_CheckAll)
	$btn_Step1_Mailboxes_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step1_Mailboxes_UncheckAll.Location = $System_Drawing_Point_Step1_UncheckAll
	$btn_Step1_Mailboxes_UncheckAll.Name = "btn_Step1_Mailboxes_UncheckAll"
	$btn_Step1_Mailboxes_UncheckAll.Size = $System_Drawing_Size_Step1_btn
	$btn_Step1_Mailboxes_UncheckAll.TabIndex = 10
	$btn_Step1_Mailboxes_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step1_Mailboxes_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step1_Mailboxes_UncheckAll.add_Click($handler_btn_Step1_Mailboxes_UncheckAll)
	$bx_Mailboxes_List.Controls.Add($btn_Step1_Mailboxes_UncheckAll)
}
#EndRegion Step1 Mailboxes tab

#Endregion "Step1 - Targets"

#Region "Step2"
$tab_Step2.BackColor = [System.Drawing.Color]::FromArgb(0,255,255,255)
$tab_Step2.Font = $font_Calibri_8pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 36
$tab_Step2.Location = $System_Drawing_Point
$tab_Step2.Name = "tab_Step2"
$tab_Step2.Padding = $System_Windows_Forms_Padding_Reusable
$tab_Step2.TabIndex = 3
$tab_Step2.Text = "  Templates  "
$tab_Step2.Size = $System_Drawing_Size_tab_1
$tab_Master.Controls.Add($tab_Step2)
$bx_Step2_Templates.Font = $font_Calibri_10pt_bold
	$System_Drawing_Point_bx_Step2 = New-Object System.Drawing.Point
	$System_Drawing_Point_bx_Step2.X = 27	# 96-69
	$System_Drawing_Point_bx_Step2.Y = 91
$bx_Step2_Templates.Location = $System_Drawing_Point_bx_Step2
$bx_Step2_Templates.Name = "bx_Step2_Templates"
	$System_Drawing_Size_bx_Step2 = New-Object System.Drawing.Size
	$System_Drawing_Size_bx_Step2.Height = 487 #482 to short
	$System_Drawing_Size_bx_Step2.Width = 536
$bx_Step2_Templates.Size = $System_Drawing_Size_bx_Step2
$bx_Step2_Templates.TabIndex = 0
$bx_Step2_Templates.TabStop = $False
$bx_Step2_Templates.Text = "Select a data collection template"
$tab_Step2.Controls.Add($bx_Step2_Templates)
$rb_Step2_Template_1.Checked = $False
$rb_Step2_Template_1.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 50
	$System_Drawing_Point.Y = 25
$rb_Step2_Template_1.Location = $System_Drawing_Point
$rb_Step2_Template_1.Name = "rb_Step2_Template_1"
$rb_Step2_Template_1.Size = $System_Drawing_Size_Reusable_chk_long
$rb_Step2_Template_1.TabIndex = 0
$rb_Step2_Template_1.Text = "Recommended tests"
$rb_Step2_Template_1.UseVisualStyleBackColor = $True
$rb_Step2_Template_1.add_Click($handler_rb_Step2_Template_1)
$bx_Step2_Templates.Controls.Add($rb_Step2_Template_1)
$rb_Step2_Template_2.Checked = $False
$rb_Step2_Template_2.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 50
	$System_Drawing_Point.Y = 50
$rb_Step2_Template_2.Location = $System_Drawing_Point
$rb_Step2_Template_2.Name = "rb_Step2_Template_2"
$rb_Step2_Template_2.Size = $System_Drawing_Size_Reusable_chk_long
$rb_Step2_Template_2.TabIndex = 0
$rb_Step2_Template_2.Text = "All tests"
$rb_Step2_Template_2.UseVisualStyleBackColor = $True
$rb_Step2_Template_2.add_Click($handler_rb_Step2_Template_2)
$bx_Step2_Templates.Controls.Add($rb_Step2_Template_2)
$rb_Step2_Template_3.Checked = $False
$rb_Step2_Template_3.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 50
	$System_Drawing_Point.Y = 75
$rb_Step2_Template_3.Location = $System_Drawing_Point
$rb_Step2_Template_3.Name = "rb_Step2_Template_3"
$rb_Step2_Template_3.Size = $System_Drawing_Size_Reusable_chk_long
$rb_Step2_Template_3.TabIndex = 0
$rb_Step2_Template_3.Text = "Minimum tests for Environmental Document"
$rb_Step2_Template_3.UseVisualStyleBackColor = $True
$rb_Step2_Template_3.add_Click($handler_rb_Step2_Template_3)
$bx_Step2_Templates.Controls.Add($rb_Step2_Template_3)
$rb_Step2_Template_4.Checked = $False
$rb_Step2_Template_4.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 50
	$System_Drawing_Point.Y = 100
$rb_Step2_Template_4.Location = $System_Drawing_Point
$rb_Step2_Template_4.Name = "rb_Step2_Template_4"
$rb_Step2_Template_4.Size = $System_Drawing_Size_Reusable_chk_long
$rb_Step2_Template_4.TabIndex = 0
$rb_Step2_Template_4.Text = "Custom Template 1"
$rb_Step2_Template_4.UseVisualStyleBackColor = $True
$rb_Step2_Template_4.add_Click($handler_rb_Step2_Template_4)
$bx_Step2_Templates.Controls.Add($rb_Step2_Template_4)
$Status_Step2.Font = $font_Calibri_10pt_normal
$Status_Step2.Location = $System_Drawing_Point_Status
$Status_Step2.Name = "Status_Step2"
$Status_Step2.Size = $System_Drawing_Size_Status
$Status_Step2.TabIndex = 12
$Status_Step2.Text = "Step 2 Status"
$tab_Step2.Controls.Add($Status_Step2)
#Endregion "Step2"

#Region "Step3 - Tests"
#Region Step3 Main
# Reusable boxes in Step3 Tabs
	$System_Drawing_Size_Step3_box = New-Object System.Drawing.Size
	$System_Drawing_Size_Step3_box.Height = 400
	$System_Drawing_Size_Step3_box.Width = 536
# Reusable check buttons in Step3 tabs
	$System_Drawing_Size_Step3_check_btn = New-Object System.Drawing.Size
	$System_Drawing_Size_Step3_check_btn.Height = 25
	$System_Drawing_Size_Step3_check_btn.Width = 150
# Reusable check/uncheck buttons in Step3 tabs
	$System_Drawing_Point_Step3_Check = New-Object System.Drawing.Point
	$System_Drawing_Point_Step3_Check.X = 50
	$System_Drawing_Point_Step3_Check.Y = 400
	$System_Drawing_Point_Step3_Uncheck = New-Object System.Drawing.Point
	$System_Drawing_Point_Step3_Uncheck.X = 300
	$System_Drawing_Point_Step3_Uncheck.Y = 400
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 36
$tab_Step3.Location = $System_Drawing_Point
$tab_Step3.Name = "tab_Step3"
$tab_Step3.Padding = $System_Windows_Forms_Padding_Reusable
$tab_Step3.TabIndex = 2
$tab_Step3.Text = "   Tests   "
$tab_Step3.Size = $System_Drawing_Size_tab_1
$tab_Master.Controls.Add($tab_Step3)
$tab_Step3_Master.Font = $font_Calibri_12pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 20
	$System_Drawing_Point.Y = 60
$tab_Step3_Master.Location = $System_Drawing_Point
$tab_Step3_Master.Name = "tab_Step3_Master"
$tab_Step3_Master.SelectedIndex = 0
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 525
	$System_Drawing_Size.Width = 550
$tab_Step3_Master.Size = $System_Drawing_Size
$tab_Step3_Master.TabIndex = 11
$tab_Step3.Controls.Add($tab_Step3_Master)
$btn_Step3_Execute.Font = $font_Calibri_14pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 20
	$System_Drawing_Point.Y = 15
$btn_Step3_Execute.Location = $System_Drawing_Point
$btn_Step3_Execute.Name = "btn_Step3_Execute"
$btn_Step3_Execute.Size = $System_Drawing_Size_buttons
$btn_Step3_Execute.TabIndex = 4
$btn_Step3_Execute.Text = "Execute"
$btn_Step3_Execute.UseVisualStyleBackColor = $True
$btn_Step3_Execute.add_Click($handler_btn_Step3_Execute_Click)
$tab_Step3.Controls.Add($btn_Step3_Execute)
$lbl_Step3_Execute.BackColor = [System.Drawing.Color]::FromArgb(0,255,255,255)
$lbl_Step3_Execute.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 138
	$System_Drawing_Point.Y = 15
$lbl_Step3_Execute.Location = $System_Drawing_Point
$lbl_Step3_Execute.Name = "lbl_Step3"
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 38
	$System_Drawing_Size.Width = 510
$lbl_Step3_Execute.Size = $System_Drawing_Size
$lbl_Step3_Execute.TabIndex = 5
$lbl_Step3_Execute.Text = "Select the functions below and click on the Execute button."
$lbl_Step3_Execute.TextAlign = 16
$tab_Step3.Controls.Add($lbl_Step3_Execute)
$status_Step3.Font = $font_Calibri_10pt_normal
$status_Step3.Location = $System_Drawing_Point_Status
$status_Step3.Name = "status_Step3"
$status_Step3.Size = $System_Drawing_Size_Status
$status_Step3.TabIndex = 10
$status_Step3.Text = "Step 3 Status"
$tab_Step3.Controls.Add($status_Step3)
#EndRegion Step3 Main

#Region Step3 Server - Tier 2
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 33
$tab_Step3_Server.Location = $System_Drawing_Point
$tab_Step3_Server.Name = "tab_Step3_Server"
$tab_Step3_Server.Padding = $System_Windows_Forms_Padding_Reusable
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 488
	$System_Drawing_Size.Width = 542
$tab_Step3_Server.Size = $System_Drawing_Size
$tab_Step3_Server.TabIndex = 0
$tab_Step3_Server.Text = "Server Functions"
$tab_Step3_Server.UseVisualStyleBackColor = $True
$tab_Step3_Master.Controls.Add($tab_Step3_Server)

# Server Tab Control
$tab_Step3_Server_Tier2.Appearance = 2
$tab_Step3_Server_Tier2.Dock = 5
$tab_Step3_Server_Tier2.Font = $font_Calibri_10pt_normal
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 32
	$System_Drawing_Size.Width = 100
$tab_Step3_Server_Tier2.ItemSize = $System_Drawing_Size
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 0
	$System_Drawing_Point.Y = 0
$tab_Step3_Server_Tier2.Location = $System_Drawing_Point
$tab_Step3_Server_Tier2.Name = "tab_Step3_Server_Tier2"
$tab_Step3_Server_Tier2.SelectedIndex = 0
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 718
	$System_Drawing_Size.Width = 665
$tab_Step3_Server_Tier2.Size = $System_Drawing_Size
$tab_Step3_Server_Tier2.TabIndex = 12
$tab_Step3_Server.Controls.Add($tab_Step3_Server_Tier2)

#EndRegion Step3 Server - Tier 2

#Region Step3 ExOrg - Tier 2
if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true) -or ($NoGUI -eq $true))
{
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = 4
		$System_Drawing_Point.Y = 33
	$tab_Step3_ExOrg.Location = $System_Drawing_Point
	$tab_Step3_ExOrg.Name = "tab_Step3_ExOrg"
	$tab_Step3_ExOrg.Padding = $System_Windows_Forms_Padding_Reusable
		$System_Drawing_Size = New-Object System.Drawing.Size
		$System_Drawing_Size.Height = 488
		$System_Drawing_Size.Width = 542
	$tab_Step3_ExOrg.Size = $System_Drawing_Size
	$tab_Step3_ExOrg.TabIndex = 0
	$tab_Step3_ExOrg.Text = "Exchange Functions"
	$tab_Step3_ExOrg.UseVisualStyleBackColor = $True
	$tab_Step3_Master.Controls.Add($tab_Step3_ExOrg)

	# ExOrg Tab Control
	$tab_Step3_ExOrg_Tier2.Appearance = 2
	$tab_Step3_ExOrg_Tier2.Dock = 5
	$tab_Step3_ExOrg_Tier2.Font = $font_Calibri_10pt_normal
		$System_Drawing_Size = New-Object System.Drawing.Size
		$System_Drawing_Size.Height = 32
		$System_Drawing_Size.Width = 100
	$tab_Step3_ExOrg_Tier2.ItemSize = $System_Drawing_Size
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = 0
		$System_Drawing_Point.Y = 0
	$tab_Step3_ExOrg_Tier2.Location = $System_Drawing_Point
	$tab_Step3_ExOrg_Tier2.Name = "tab_Step3_ExOrg_Tier2"
	$tab_Step3_ExOrg_Tier2.SelectedIndex = 0
		$System_Drawing_Size = New-Object System.Drawing.Size
		$System_Drawing_Size.Height = 718
		$System_Drawing_Size.Width = 665
	$tab_Step3_ExOrg_Tier2.Size = $System_Drawing_Size
	$tab_Step3_ExOrg_Tier2.TabIndex = 12
	$tab_Step3_ExOrg.Controls.Add($tab_Step3_ExOrg_Tier2)
}
#EndRegion Step3 ExOrg - Tier 2

#Region Step3 Domain Controller tab
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 33
$tab_Step3_DC.Location = $System_Drawing_Point
$tab_Step3_DC.Name = "tab_Step3_DC"
$tab_Step3_DC.Padding = $System_Windows_Forms_Padding_Reusable
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 488
	$System_Drawing_Size.Width = 542
$tab_Step3_DC.Size = $System_Drawing_Size
$tab_Step3_DC.TabIndex = 0
$tab_Step3_DC.Text = "Domain Controllers"
$tab_Step3_DC.UseVisualStyleBackColor = $True
$tab_Step3_Server_Tier2.Controls.Add($tab_Step3_DC)
$bx_DC_Functions.Dock = 5
$bx_DC_Functions.Font = $font_Calibri_10pt_bold
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 3
	$System_Drawing_Point.Y = 3
$bx_DC_Functions.Location = $System_Drawing_Point
$bx_DC_Functions.Name = "bx_DC_Functions"
$bx_DC_Functions.Size = $System_Drawing_Size_Step3_box
$bx_DC_Functions.TabIndex = 7
$bx_DC_Functions.TabStop = $False
$tab_Step3_DC.Controls.Add($bx_DC_Functions)
	$Col_1_loc = 35
	$Col_2_loc = 290
	$Row_1_loc = 25
	$Row_2_loc = 25
$chk_DC_Win32_Bios.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
$chk_DC_Win32_Bios.Location = $System_Drawing_Point
$chk_DC_Win32_Bios.Name = "chk_DC_Win32_Bios"
$chk_DC_Win32_Bios.Size = $System_Drawing_Size_Reusable_chk
$chk_DC_Win32_Bios.TabIndex = 0
$chk_DC_Win32_Bios.Text = "Win32_Bios"
$chk_DC_Win32_Bios.UseVisualStyleBackColor = $True
$bx_DC_Functions.Controls.Add($chk_DC_Win32_Bios)
$chk_DC_Win32_ComputerSystem.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
$chk_DC_Win32_ComputerSystem.Location = $System_Drawing_Point
$chk_DC_Win32_ComputerSystem.Name = "chk_DC_Win32_ComputerSystem"
$chk_DC_Win32_ComputerSystem.Size = $System_Drawing_Size_Reusable_chk
$chk_DC_Win32_ComputerSystem.TabIndex = 0
$chk_DC_Win32_ComputerSystem.Text = "Win32_ComputerSystem"
$chk_DC_Win32_ComputerSystem.UseVisualStyleBackColor = $True
$bx_DC_Functions.Controls.Add($chk_DC_Win32_ComputerSystem)
$chk_DC_Win32_LogicalDisk.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
$chk_DC_Win32_LogicalDisk.Location = $System_Drawing_Point
$chk_DC_Win32_LogicalDisk.Name = "chk_DC_Win32_LogicalDisk"
$chk_DC_Win32_LogicalDisk.Size = $System_Drawing_Size_Reusable_chk
$chk_DC_Win32_LogicalDisk.TabIndex = 1
$chk_DC_Win32_LogicalDisk.Text = "Win32_LogicalDisk"
$chk_DC_Win32_LogicalDisk.UseVisualStyleBackColor = $True
$bx_DC_Functions.Controls.Add($chk_DC_Win32_LogicalDisk)
$chk_DC_Win32_NetworkAdapter.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
$chk_DC_Win32_NetworkAdapter.Location = $System_Drawing_Point
$chk_DC_Win32_NetworkAdapter.Name = "chk_DC_Win32_NetworkAdapter"
$chk_DC_Win32_NetworkAdapter.Size = $System_Drawing_Size_Reusable_chk
$chk_DC_Win32_NetworkAdapter.TabIndex = 1
$chk_DC_Win32_NetworkAdapter.Text = "Win32_NetworkAdapter"
$chk_DC_Win32_NetworkAdapter.UseVisualStyleBackColor = $True
$bx_DC_Functions.Controls.Add($chk_DC_Win32_NetworkAdapter)
$chk_DC_Win32_NetworkAdapterConfig.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
$chk_DC_Win32_NetworkAdapterConfig.Location = $System_Drawing_Point
$chk_DC_Win32_NetworkAdapterConfig.Name = "chk_DC_Win32_NetworkAdapterConfig"
$chk_DC_Win32_NetworkAdapterConfig.Size = $System_Drawing_Size_Reusable_chk
$chk_DC_Win32_NetworkAdapterConfig.TabIndex = 2
$chk_DC_Win32_NetworkAdapterConfig.Text = "Win32_NetworkAdapterConfig"
$chk_DC_Win32_NetworkAdapterConfig.UseVisualStyleBackColor = $True
$bx_DC_Functions.Controls.Add($chk_DC_Win32_NetworkAdapterConfig)
$chk_DC_Win32_OperatingSystem.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
$chk_DC_Win32_OperatingSystem.Location = $System_Drawing_Point
$chk_DC_Win32_OperatingSystem.Name = "chk_DC_Win32_OperatingSystem"
$chk_DC_Win32_OperatingSystem.Size = $System_Drawing_Size_Reusable_chk
$chk_DC_Win32_OperatingSystem.TabIndex = 3
$chk_DC_Win32_OperatingSystem.Text = "Win32_OperatingSystem"
$chk_DC_Win32_OperatingSystem.UseVisualStyleBackColor = $True
$bx_DC_Functions.Controls.Add($chk_DC_Win32_OperatingSystem)
$chk_DC_Win32_PageFileUsage.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
$chk_DC_Win32_PageFileUsage.Location = $System_Drawing_Point
$chk_DC_Win32_PageFileUsage.Name = "chk_DC_Win32_PageFileUsage"
$chk_DC_Win32_PageFileUsage.Size = $System_Drawing_Size_Reusable_chk
$chk_DC_Win32_PageFileUsage.TabIndex = 4
$chk_DC_Win32_PageFileUsage.Text = "Win32_PageFileUsage"
$chk_DC_Win32_PageFileUsage.UseVisualStyleBackColor = $True
$bx_DC_Functions.Controls.Add($chk_DC_Win32_PageFileUsage)
$chk_DC_Win32_PhysicalMemory.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
$chk_DC_Win32_PhysicalMemory.Location = $System_Drawing_Point
$chk_DC_Win32_PhysicalMemory.Name = "chk_DC_Win32_PhysicalMemory"
$chk_DC_Win32_PhysicalMemory.Size = $System_Drawing_Size_Reusable_chk
$chk_DC_Win32_PhysicalMemory.TabIndex = 5
$chk_DC_Win32_PhysicalMemory.Text = "Win32_PhysicalMemory"
$chk_DC_Win32_PhysicalMemory.UseVisualStyleBackColor = $True
$bx_DC_Functions.Controls.Add($chk_DC_Win32_PhysicalMemory)
$chk_DC_Win32_Processor.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
$chk_DC_Win32_Processor.Location = $System_Drawing_Point
$chk_DC_Win32_Processor.Name = "chk_DC_Win32_Processor"
$chk_DC_Win32_Processor.Size = $System_Drawing_Size_Reusable_chk
$chk_DC_Win32_Processor.TabIndex = 6
$chk_DC_Win32_Processor.Text = "Win32_Processor"
$chk_DC_Win32_Processor.UseVisualStyleBackColor = $True
$bx_DC_Functions.Controls.Add($chk_DC_Win32_Processor)
$chk_DC_Registry_AD.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
$chk_DC_Registry_AD.Location = $System_Drawing_Point
$chk_DC_Registry_AD.Name = "chk_DC_Registry_AD"
$chk_DC_Registry_AD.Size = $System_Drawing_Size_Reusable_chk
$chk_DC_Registry_AD.TabIndex = 7
$chk_DC_Registry_AD.Text = "Registry - AD"
$chk_DC_Registry_AD.UseVisualStyleBackColor = $True
$bx_DC_Functions.Controls.Add($chk_DC_Registry_AD)
$chk_DC_Registry_OS.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
$chk_DC_Registry_OS.Location = $System_Drawing_Point
$chk_DC_Registry_OS.Name = "chk_DC_Registry_OS"
$chk_DC_Registry_OS.Size = $System_Drawing_Size_Reusable_chk
$chk_DC_Registry_OS.TabIndex = 8
$chk_DC_Registry_OS.Text = "Registry - OS"
$chk_DC_Registry_OS.UseVisualStyleBackColor = $True
$bx_DC_Functions.Controls.Add($chk_DC_Registry_OS)
$chk_DC_Registry_Software.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
$chk_DC_Registry_Software.Location = $System_Drawing_Point
$chk_DC_Registry_Software.Name = "chk_DC_Registry_Software"
$chk_DC_Registry_Software.Size = $System_Drawing_Size_Reusable_chk
$chk_DC_Registry_Software.TabIndex = 9
$chk_DC_Registry_Software.Text = "Registry - Software"
$chk_DC_Registry_Software.UseVisualStyleBackColor = $True
$bx_DC_Functions.Controls.Add($chk_DC_Registry_Software)
$chk_DC_MicrosoftDNS_Zone.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_2_loc
		$System_Drawing_Point.Y = $Row_2_loc
		$Row_2_loc += 25
$chk_DC_MicrosoftDNS_Zone.Location = $System_Drawing_Point
$chk_DC_MicrosoftDNS_Zone.Name = "chk_DC_MicrosoftDNS_Zone"
$chk_DC_MicrosoftDNS_Zone.Size = $System_Drawing_Size_Reusable_chk
$chk_DC_MicrosoftDNS_Zone.TabIndex = 10
$chk_DC_MicrosoftDNS_Zone.Text = "MicrosoftDNS_Zone"
$chk_DC_MicrosoftDNS_Zone.UseVisualStyleBackColor = $True
$bx_DC_Functions.Controls.Add($chk_DC_MicrosoftDNS_Zone)
$chk_DC_MSAD_DomainController.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_2_loc
		$System_Drawing_Point.Y = $Row_2_loc
		$Row_2_loc += 25
$chk_DC_MSAD_DomainController.Location = $System_Drawing_Point
$chk_DC_MSAD_DomainController.Name = "chk_DC_MSAD_DomainController"
$chk_DC_MSAD_DomainController.Size = $System_Drawing_Size_Reusable_chk
$chk_DC_MSAD_DomainController.TabIndex = 10
$chk_DC_MSAD_DomainController.Text = "MSAD_DomainController"
$chk_DC_MSAD_DomainController.UseVisualStyleBackColor = $True
$bx_DC_Functions.Controls.Add($chk_DC_MSAD_DomainController)
$chk_DC_MSAD_ReplNeighbor.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_2_loc
		$System_Drawing_Point.Y = $Row_2_loc
		$Row_2_loc += 25
$chk_DC_MSAD_ReplNeighbor.Location = $System_Drawing_Point
$chk_DC_MSAD_ReplNeighbor.Name = "chk_DC_MSAD_ReplNeighbor"
$chk_DC_MSAD_ReplNeighbor.Size = $System_Drawing_Size_Reusable_chk
$chk_DC_MSAD_ReplNeighbor.TabIndex = 10
$chk_DC_MSAD_ReplNeighbor.Text = "MSAD_ReplNeighbor"
$chk_DC_MSAD_ReplNeighbor.UseVisualStyleBackColor = $True
$bx_DC_Functions.Controls.Add($chk_DC_MSAD_ReplNeighbor)
$btn_Step3_DC_CheckAll.Font = $font_Calibri_10pt_normal
$btn_Step3_DC_CheckAll.Location = $System_Drawing_Point_Step3_Check
$btn_Step3_DC_CheckAll.Name = "btn_Step3_DC_CheckAll"
$btn_Step3_DC_CheckAll.Size = $System_Drawing_Size_Step3_check_btn
$btn_Step3_DC_CheckAll.TabIndex = 10
$btn_Step3_DC_CheckAll.Text = "Check all on this tab"
$btn_Step3_DC_CheckAll.UseVisualStyleBackColor = $True
$btn_Step3_DC_CheckAll.add_Click($handler_btn_Step3_DC_CheckAll_Click)
$bx_DC_Functions.Controls.Add($btn_Step3_DC_CheckAll)
$btn_Step3_DC_UncheckAll.Font = $font_Calibri_10pt_normal
$btn_Step3_DC_UncheckAll.Location = $System_Drawing_Point_Step3_Uncheck
$btn_Step3_DC_UncheckAll.Name = "btn_Step3_DC_UncheckAll"
$btn_Step3_DC_UncheckAll.Size = $System_Drawing_Size_Step3_check_btn
$btn_Step3_DC_UncheckAll.TabIndex = 11
$btn_Step3_DC_UncheckAll.Text = "Uncheck all on this tab"
$btn_Step3_DC_UncheckAll.UseVisualStyleBackColor = $True
$btn_Step3_DC_UncheckAll.add_Click($handler_btn_Step3_DC_UncheckAll_Click)
$bx_DC_Functions.Controls.Add($btn_Step3_DC_UncheckAll)
#EndRegion Step3 Domain Controller tab

#Region Step3 Exchange Servers tab
$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 33
$tab_Step3_Exchange.Location = $System_Drawing_Point
$tab_Step3_Exchange.Name = "tab_Step3_Exchange"
$tab_Step3_Exchange.Padding = $System_Windows_Forms_Padding_Reusable
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 488
	$System_Drawing_Size.Width = 542
$tab_Step3_Exchange.Size = $System_Drawing_Size
$tab_Step3_Exchange.TabIndex = 1
$tab_Step3_Exchange.Text = "Exchange Servers"
$tab_Step3_Exchange.UseVisualStyleBackColor = $True
$tab_Step3_Server_Tier2.Controls.Add($tab_Step3_Exchange)
$bx_Exchange_Functions.Dock = 5
$bx_Exchange_Functions.Font = $font_Calibri_10pt_bold
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 3
	$System_Drawing_Point.Y = 3
$bx_Exchange_Functions.Location = $System_Drawing_Point
$bx_Exchange_Functions.Name = "bx_Exchange_Functions"
$bx_Exchange_Functions.Size = $System_Drawing_Size_Step3_box
$bx_Exchange_Functions.TabIndex = 8
$bx_Exchange_Functions.TabStop = $False
$tab_Step3_Exchange.Controls.Add($bx_Exchange_Functions)
	$Col_1_loc = 35
	$Col_2_loc = 290
	$Row_1_loc = 25
	$Row_2_loc = 25
$chk_Ex_Win32_Bios.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
$chk_Ex_Win32_Bios.Location = $System_Drawing_Point
$chk_Ex_Win32_Bios.Name = "chk_Ex_Win32_Bios"
$chk_Ex_Win32_Bios.Size = $System_Drawing_Size_Reusable_chk
$chk_Ex_Win32_Bios.TabIndex = 0
$chk_Ex_Win32_Bios.Text = "Win32_Bios"
$chk_Ex_Win32_Bios.UseVisualStyleBackColor = $True
$bx_Exchange_Functions.Controls.Add($chk_Ex_Win32_Bios)
$chk_Ex_Win32_ComputerSystem.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
$chk_Ex_Win32_ComputerSystem.Location = $System_Drawing_Point
$chk_Ex_Win32_ComputerSystem.Name = "chk_Ex_Win32_ComputerSystem"
$chk_Ex_Win32_ComputerSystem.Size = $System_Drawing_Size_Reusable_chk
$chk_Ex_Win32_ComputerSystem.TabIndex = 0
$chk_Ex_Win32_ComputerSystem.Text = "Win32_ComputerSystem"
$chk_Ex_Win32_ComputerSystem.UseVisualStyleBackColor = $True
$bx_Exchange_Functions.Controls.Add($chk_Ex_Win32_ComputerSystem)
$chk_Ex_Win32_LogicalDisk.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
$chk_Ex_Win32_LogicalDisk.Location = $System_Drawing_Point
$chk_Ex_Win32_LogicalDisk.Name = "chk_Ex_Win32_LogicalDisk"
$chk_Ex_Win32_LogicalDisk.Size = $System_Drawing_Size_Reusable_chk
$chk_Ex_Win32_LogicalDisk.TabIndex = 1
$chk_Ex_Win32_LogicalDisk.Text = "Win32_LogicalDisk"
$chk_Ex_Win32_LogicalDisk.UseVisualStyleBackColor = $True
$bx_Exchange_Functions.Controls.Add($chk_Ex_Win32_LogicalDisk)
$chk_Ex_Win32_NetworkAdapter.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
$chk_Ex_Win32_NetworkAdapter.Location = $System_Drawing_Point
$chk_Ex_Win32_NetworkAdapter.Name = "chk_Ex_Win32_NetworkAdapter"
$chk_Ex_Win32_NetworkAdapter.Size = $System_Drawing_Size_Reusable_chk
$chk_Ex_Win32_NetworkAdapter.TabIndex = 1
$chk_Ex_Win32_NetworkAdapter.Text = "Win32_NetworkAdapter"
$chk_Ex_Win32_NetworkAdapter.UseVisualStyleBackColor = $True
$bx_Exchange_Functions.Controls.Add($chk_Ex_Win32_NetworkAdapter)
$chk_Ex_Win32_NetworkAdapterConfig.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
$chk_Ex_Win32_NetworkAdapterConfig.Location = $System_Drawing_Point
$chk_Ex_Win32_NetworkAdapterConfig.Name = "chk_Ex_Win32_NetworkAdapterConfig"
$chk_Ex_Win32_NetworkAdapterConfig.Size = $System_Drawing_Size_Reusable_chk
$chk_Ex_Win32_NetworkAdapterConfig.TabIndex = 2
$chk_Ex_Win32_NetworkAdapterConfig.Text = "Win32_NetworkAdapterConfig"
$chk_Ex_Win32_NetworkAdapterConfig.UseVisualStyleBackColor = $True
$bx_Exchange_Functions.Controls.Add($chk_Ex_Win32_NetworkAdapterConfig)
$chk_Ex_Win32_OperatingSystem.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
$chk_Ex_Win32_OperatingSystem.Location = $System_Drawing_Point
$chk_Ex_Win32_OperatingSystem.Name = "chk_Ex_Win32_OperatingSystem"
$chk_Ex_Win32_OperatingSystem.Size = $System_Drawing_Size_Reusable_chk
$chk_Ex_Win32_OperatingSystem.TabIndex = 3
$chk_Ex_Win32_OperatingSystem.Text = "Win32_OperatingSystem"
$chk_Ex_Win32_OperatingSystem.UseVisualStyleBackColor = $True
$bx_Exchange_Functions.Controls.Add($chk_Ex_Win32_OperatingSystem)
$chk_Ex_Win32_PageFileUsage.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
$chk_Ex_Win32_PageFileUsage.Location = $System_Drawing_Point
$chk_Ex_Win32_PageFileUsage.Name = "chk_Ex_Win32_PageFileUsage"
$chk_Ex_Win32_PageFileUsage.Size = $System_Drawing_Size_Reusable_chk
$chk_Ex_Win32_PageFileUsage.TabIndex = 4
$chk_Ex_Win32_PageFileUsage.Text = "Win32_PageFileUsage"
$chk_Ex_Win32_PageFileUsage.UseVisualStyleBackColor = $True
$bx_Exchange_Functions.Controls.Add($chk_Ex_Win32_PageFileUsage)
$chk_Ex_Win32_PhysicalMemory.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
$chk_Ex_Win32_PhysicalMemory.Location = $System_Drawing_Point
$chk_Ex_Win32_PhysicalMemory.Name = "chk_Ex_Win32_PhysicalMemory"
$chk_Ex_Win32_PhysicalMemory.Size = $System_Drawing_Size_Reusable_chk
$chk_Ex_Win32_PhysicalMemory.TabIndex = 5
$chk_Ex_Win32_PhysicalMemory.Text = "Win32_PhysicalMemory"
$chk_Ex_Win32_PhysicalMemory.UseVisualStyleBackColor = $True
$bx_Exchange_Functions.Controls.Add($chk_Ex_Win32_PhysicalMemory)
$chk_Ex_Win32_Processor.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
$chk_Ex_Win32_Processor.Location = $System_Drawing_Point
$chk_Ex_Win32_Processor.Name = "chk_Ex_Win32_Processor"
$chk_Ex_Win32_Processor.Size = $System_Drawing_Size_Reusable_chk
$chk_Ex_Win32_Processor.TabIndex = 6
$chk_Ex_Win32_Processor.Text = "Win32_Processor"
$chk_Ex_Win32_Processor.UseVisualStyleBackColor = $True
$bx_Exchange_Functions.Controls.Add($chk_Ex_Win32_Processor)
$chk_Ex_Registry_Ex.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
$chk_Ex_Registry_Ex.Location = $System_Drawing_Point
$chk_Ex_Registry_Ex.Name = "chk_Ex_Registry_Ex"
$chk_Ex_Registry_Ex.Size = $System_Drawing_Size_Reusable_chk
$chk_Ex_Registry_Ex.TabIndex = 7
$chk_Ex_Registry_Ex.Text = "Registry - Exchange"
$chk_Ex_Registry_Ex.UseVisualStyleBackColor = $True
$bx_Exchange_Functions.Controls.Add($chk_Ex_Registry_Ex)
$chk_Ex_Registry_OS.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
$chk_Ex_Registry_OS.Location = $System_Drawing_Point
$chk_Ex_Registry_OS.Name = "chk_Ex_Registry_OS"
$chk_Ex_Registry_OS.Size = $System_Drawing_Size_Reusable_chk
$chk_Ex_Registry_OS.TabIndex = 8
$chk_Ex_Registry_OS.Text = "Registry - OS"
$chk_Ex_Registry_OS.UseVisualStyleBackColor = $True
$bx_Exchange_Functions.Controls.Add($chk_Ex_Registry_OS)
$chk_Ex_Registry_Software.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
$chk_Ex_Registry_Software.Location = $System_Drawing_Point
$chk_Ex_Registry_Software.Name = "chk_Ex_Registry_Software"
$chk_Ex_Registry_Software.Size = $System_Drawing_Size_Reusable_chk
$chk_Ex_Registry_Software.TabIndex = 9
$chk_Ex_Registry_Software.Text = "Registry - Software"
$chk_Ex_Registry_Software.UseVisualStyleBackColor = $True
$bx_Exchange_Functions.Controls.Add($chk_Ex_Registry_Software)
$btn_Step3_Ex_CheckAll.Font = $font_Calibri_10pt_normal
$btn_Step3_Ex_CheckAll.Location = $System_Drawing_Point_Step3_Check
$btn_Step3_Ex_CheckAll.Name = "btn_Step3_Ex_CheckAll"
$btn_Step3_Ex_CheckAll.Size = $System_Drawing_Size_Step3_check_btn
$btn_Step3_Ex_CheckAll.TabIndex = 10
$btn_Step3_Ex_CheckAll.Text = "Check all on this tab"
$btn_Step3_Ex_CheckAll.UseVisualStyleBackColor = $True
$btn_Step3_Ex_CheckAll.add_Click($handler_btn_Step3_Ex_CheckAll_Click)
$bx_Exchange_Functions.Controls.Add($btn_Step3_Ex_CheckAll)
$btn_Step3_Ex_UncheckAll.Font = $font_Calibri_10pt_normal
$btn_Step3_Ex_UncheckAll.Location = $System_Drawing_Point_Step3_Uncheck
$btn_Step3_Ex_UncheckAll.Name = "btn_Step3_Ex_UncheckAll"
$btn_Step3_Ex_UncheckAll.Size = $System_Drawing_Size_Step3_check_btn
$btn_Step3_Ex_UncheckAll.TabIndex = 11
$btn_Step3_Ex_UncheckAll.Text = "Uncheck all on this tab"
$btn_Step3_Ex_UncheckAll.UseVisualStyleBackColor = $True
$btn_Step3_Ex_UncheckAll.add_Click($handler_btn_Step3_Ex_UncheckAll_Click)
$bx_Exchange_Functions.Controls.Add($btn_Step3_Ex_UncheckAll)
#EndRegion Step3 Exchange Servers tab

#Region Step3 Cluster Nodes tab
if (($Exchange2007Powershell -eq $true) -or ($Exchange2003orEarlier -eq $true) -or ($NoEMS -eq $true) -or ($NoGUI -eq $true))
{
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = 4
		$System_Drawing_Point.Y = 33
	$tab_Step3_Cluster.Location = $System_Drawing_Point
	$tab_Step3_Cluster.Name = "tab_Step3_Cluster"
	$tab_Step3_Cluster.Padding = $System_Windows_Forms_Padding_Reusable
		$System_Drawing_Size = New-Object System.Drawing.Size
		$System_Drawing_Size.Height = 488
		$System_Drawing_Size.Width = 542
	$tab_Step3_Cluster.Size = $System_Drawing_Size
	$tab_Step3_Cluster.TabIndex = 2
	$tab_Step3_Cluster.Text = "Cluster Nodes"
	$tab_Step3_Cluster.UseVisualStyleBackColor = $True
	$tab_Step3_Server_Tier2.Controls.Add($tab_Step3_Cluster)
	$bx_Cluster_Functions.Dock = 5
	$bx_Cluster_Functions.Font = $font_Calibri_10pt_bold
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = 3
		$System_Drawing_Point.Y = 3
	$bx_Cluster_Functions.Location = $System_Drawing_Point
	$bx_Cluster_Functions.Name = "bx_Cluster_Functions"
	$bx_Cluster_Functions.Size = $System_Drawing_Size_Step3_box
	$bx_Cluster_Functions.TabIndex = 12
	$bx_Cluster_Functions.TabStop = $False
	$tab_Step3_Cluster.Controls.Add($bx_Cluster_Functions)
	$Col_1_loc = 35
	$Col_2_loc = 290
	$Row_1_loc = 25
	$Row_2_loc = 25
	$chk_Cluster_MSCluster_Node.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Cluster_MSCluster_Node.Location = $System_Drawing_Point
	$chk_Cluster_MSCluster_Node.Name = "chk_Cluster_MSCluster_Node"
	$chk_Cluster_MSCluster_Node.Size = $System_Drawing_Size_Reusable_chk
	$chk_Cluster_MSCluster_Node.TabIndex = 0
	$chk_Cluster_MSCluster_Node.Text = "MSCluster_Node"
	$chk_Cluster_MSCluster_Node.UseVisualStyleBackColor = $True
	$bx_Cluster_Functions.Controls.Add($chk_Cluster_MSCluster_Node)
	$chk_Cluster_MSCluster_Network.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Cluster_MSCluster_Network.Location = $System_Drawing_Point
	$chk_Cluster_MSCluster_Network.Name = "chk_Cluster_MSCluster_Network"
	$chk_Cluster_MSCluster_Network.Size = $System_Drawing_Size_Reusable_chk
	$chk_Cluster_MSCluster_Network.TabIndex = 2
	$chk_Cluster_MSCluster_Network.Text = "MSCluster_Network"
	$chk_Cluster_MSCluster_Network.UseVisualStyleBackColor = $True
	$bx_Cluster_Functions.Controls.Add($chk_Cluster_MSCluster_Network)
	$chk_Cluster_MSCluster_Resource.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Cluster_MSCluster_Resource.Location = $System_Drawing_Point
	$chk_Cluster_MSCluster_Resource.Name = "chk_Cluster_MSCluster_Resource"
	$chk_Cluster_MSCluster_Resource.Size = $System_Drawing_Size_Reusable_chk
	$chk_Cluster_MSCluster_Resource.TabIndex = 3
	$chk_Cluster_MSCluster_Resource.Text = "MSCluster_Resource"
	$chk_Cluster_MSCluster_Resource.UseVisualStyleBackColor = $True
	$bx_Cluster_Functions.Controls.Add($chk_Cluster_MSCluster_Resource)
	$chk_Cluster_MSCluster_ResourceGroup.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Cluster_MSCluster_ResourceGroup.Location = $System_Drawing_Point
	$chk_Cluster_MSCluster_ResourceGroup.Name = "chk_Cluster_MSCluster_ResourceGroup"
	$chk_Cluster_MSCluster_ResourceGroup.Size = $System_Drawing_Size_Reusable_chk
	$chk_Cluster_MSCluster_ResourceGroup.TabIndex = 4
	$chk_Cluster_MSCluster_ResourceGroup.Text = "MSCluster_ResourceGroup"
	$chk_Cluster_MSCluster_ResourceGroup.UseVisualStyleBackColor = $True
	$bx_Cluster_Functions.Controls.Add($chk_Cluster_MSCluster_ResourceGroup)
	$btn_Step3_Cluster_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Cluster_CheckAll.Location = $System_Drawing_Point_Step3_Check
	$btn_Step3_Cluster_CheckAll.Name = "btn_Step3_Cluster_CheckAll"
	$btn_Step3_Cluster_CheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_Cluster_CheckAll.TabIndex = 16
	$btn_Step3_Cluster_CheckAll.Text = "Check all on this tab"
	$btn_Step3_Cluster_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Cluster_CheckAll.add_Click($handler_btn_Step3_Cluster_CheckAll_Click)
	$bx_Cluster_Functions.Controls.Add($btn_Step3_Cluster_CheckAll)
	$btn_Step3_Cluster_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Cluster_UncheckAll.Location = $System_Drawing_Point_Step3_Uncheck
	$btn_Step3_Cluster_UncheckAll.Name = "btn_Step3_Cluster_UncheckAll"
	$btn_Step3_Cluster_UncheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_Cluster_UncheckAll.TabIndex = 17
	$btn_Step3_Cluster_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step3_Cluster_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Cluster_UncheckAll.add_Click($handler_btn_Step3_Cluster_UncheckAll_Click)
	$bx_Cluster_Functions.Controls.Add($btn_Step3_Cluster_UncheckAll)
}
#EndRegion Step3 Cluster Nodes tab

#Region Step3 Client Access tab
if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true) -or ($NoGUI -eq $true))
{
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = 4
		$System_Drawing_Point.Y = 33
	$tab_Step3_ClientAccess.Location = $System_Drawing_Point
	$tab_Step3_ClientAccess.Name = "tab_Step3_ClientAccess"
	$tab_Step3_ClientAccess.Padding = $System_Windows_Forms_Padding_Reusable
		$System_Drawing_Size = New-Object System.Drawing.Size
		$System_Drawing_Size.Height = 488
		$System_Drawing_Size.Width = 542
	$tab_Step3_ClientAccess.Size = $System_Drawing_Size
	$tab_Step3_ClientAccess.TabIndex = 3
	$tab_Step3_ClientAccess.Text = "Client Access"
	$tab_Step3_ClientAccess.UseVisualStyleBackColor = $True
	$tab_Step3_ExOrg_Tier2.Controls.Add($tab_Step3_ClientAccess)
	$bx_ClientAccess_Functions.Dock = 5
	$bx_ClientAccess_Functions.Font = $font_Calibri_10pt_bold
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = 3
		$System_Drawing_Point.Y = 3
	$bx_ClientAccess_Functions.Location = $System_Drawing_Point
	$bx_ClientAccess_Functions.Name = "bx_ClientAccess_Functions"
	$bx_ClientAccess_Functions.Size = $System_Drawing_Size_Step3_box
	$bx_ClientAccess_Functions.TabIndex = 9
	$bx_ClientAccess_Functions.TabStop = $False
	$tab_Step3_ClientAccess.Controls.Add($bx_ClientAccess_Functions)
	$btn_Step3_ClientAccess_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_ClientAccess_CheckAll.Location = $System_Drawing_Point_Step3_Check
	$btn_Step3_ClientAccess_CheckAll.Name = "btn_Step3_ClientAccess_CheckAll"
	$btn_Step3_ClientAccess_CheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_ClientAccess_CheckAll.TabIndex = 28
	$btn_Step3_ClientAccess_CheckAll.Text = "Check all on this tab"
	$btn_Step3_ClientAccess_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_ClientAccess_CheckAll.add_Click($handler_btn_Step3_ClientAccess_CheckAll_Click)
	$bx_ClientAccess_Functions.Controls.Add($btn_Step3_ClientAccess_CheckAll)
	$btn_Step3_ClientAccess_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_ClientAccess_UncheckAll.Location = $System_Drawing_Point_Step3_Uncheck
	$btn_Step3_ClientAccess_UncheckAll.Name = "btn_Step3_ClientAccess_UncheckAll"
	$btn_Step3_ClientAccess_UncheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_ClientAccess_UncheckAll.TabIndex = 29
	$btn_Step3_ClientAccess_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step3_ClientAccess_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_ClientAccess_UncheckAll.add_Click($handler_btn_Step3_ClientAccess_UncheckAll_Click)
	$bx_ClientAccess_Functions.Controls.Add($btn_Step3_ClientAccess_UncheckAll)
	$Col_1_loc = 35
	$Col_2_loc = 290
	$Row_1_loc = 25
	$Row_2_loc = 25
	$chk_Org_Get_ActiveSyncDevice.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_ActiveSyncDevice.Location = $System_Drawing_Point
	$chk_Org_Get_ActiveSyncDevice.Name = "chk_Org_Get_ActiveSyncDevice"
	$chk_Org_Get_ActiveSyncDevice.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_ActiveSyncDevice.TabIndex = 0
	$chk_Org_Get_ActiveSyncDevice.Text = "Get-ActiveSyncDevice (E14)"
	$chk_Org_Get_ActiveSyncDevice.UseVisualStyleBackColor = $True
	$bx_ClientAccess_Functions.Controls.Add($chk_Org_Get_ActiveSyncDevice)
	$chk_Org_Get_ActiveSyncPolicy.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_ActiveSyncPolicy.Location = $System_Drawing_Point
	$chk_Org_Get_ActiveSyncPolicy.Name = "chk_Org_Get_ActiveSyncPolicy"
	$chk_Org_Get_ActiveSyncPolicy.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_ActiveSyncPolicy.TabIndex = 2
	$chk_Org_Get_ActiveSyncPolicy.Text = "Get-ActiveSyncMailboxPolicy"
	$chk_Org_Get_ActiveSyncPolicy.UseVisualStyleBackColor = $True
	$bx_ClientAccess_Functions.Controls.Add($chk_Org_Get_ActiveSyncPolicy)
	$chk_Org_Get_ActiveSyncVirtualDirectory.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_ActiveSyncVirtualDirectory.Location = $System_Drawing_Point
	$chk_Org_Get_ActiveSyncVirtualDirectory.Name = "chk_Org_Get_ActiveSyncVirtualDirectory"
	$chk_Org_Get_ActiveSyncVirtualDirectory.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_ActiveSyncVirtualDirectory.TabIndex = 3
	$chk_Org_Get_ActiveSyncVirtualDirectory.Text = "Get-ActiveSyncVirtualDirectory"
	$chk_Org_Get_ActiveSyncVirtualDirectory.UseVisualStyleBackColor = $True
	$bx_ClientAccess_Functions.Controls.Add($chk_Org_Get_ActiveSyncVirtualDirectory)
	$chk_Org_Get_AutodiscoverVirtualDirectory.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_AutodiscoverVirtualDirectory.Location = $System_Drawing_Point
	$chk_Org_Get_AutodiscoverVirtualDirectory.Name = "chk_Org_Get_AutodiscoverVirtualDirectory"
	$chk_Org_Get_AutodiscoverVirtualDirectory.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_AutodiscoverVirtualDirectory.TabIndex = 7
	$chk_Org_Get_AutodiscoverVirtualDirectory.Text = "Get-AutodiscoverVirtualDirectory"
	$chk_Org_Get_AutodiscoverVirtualDirectory.UseVisualStyleBackColor = $True
	$bx_ClientAccess_Functions.Controls.Add($chk_Org_Get_AutodiscoverVirtualDirectory)
	$chk_Org_Get_AvailabilityAddressSpace.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_AvailabilityAddressSpace.Location = $System_Drawing_Point
	$chk_Org_Get_AvailabilityAddressSpace.Name = "chk_Org_Get_AvailabilityAddressSpace"
	$chk_Org_Get_AvailabilityAddressSpace.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_AvailabilityAddressSpace.TabIndex = 10
	$chk_Org_Get_AvailabilityAddressSpace.Text = "Get-AvailabilityAddressSpace (E14)"
	$chk_Org_Get_AvailabilityAddressSpace.UseVisualStyleBackColor = $True
	$bx_ClientAccess_Functions.Controls.Add($chk_Org_Get_AvailabilityAddressSpace)
	$chk_Org_Get_ClientAccessArray.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_ClientAccessArray.Location = $System_Drawing_Point
	$chk_Org_Get_ClientAccessArray.Name = "chk_Org_Get_ClientAccessArray"
	$chk_Org_Get_ClientAccessArray.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_ClientAccessArray.TabIndex = 10
	$chk_Org_Get_ClientAccessArray.Text = "Get-ClientAccessArray (E14)"
	$chk_Org_Get_ClientAccessArray.UseVisualStyleBackColor = $True
	$bx_ClientAccess_Functions.Controls.Add($chk_Org_Get_ClientAccessArray)
	$chk_Org_Get_ClientAccessServer.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_ClientAccessServer.Location = $System_Drawing_Point
	$chk_Org_Get_ClientAccessServer.Name = "chk_Org_Get_ClientAccessServer"
	$chk_Org_Get_ClientAccessServer.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_ClientAccessServer.TabIndex = 11
	$chk_Org_Get_ClientAccessServer.Text = "Get-ClientAccessServer"
	$chk_Org_Get_ClientAccessServer.UseVisualStyleBackColor = $True
	$bx_ClientAccess_Functions.Controls.Add($chk_Org_Get_ClientAccessServer)
	$chk_Org_Get_ECPVirtualDirectory.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_ECPVirtualDirectory.Location = $System_Drawing_Point
	$chk_Org_Get_ECPVirtualDirectory.Name = "chk_Org_Get_ECPVirtualDirectory"
	$chk_Org_Get_ECPVirtualDirectory.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_ECPVirtualDirectory.TabIndex = 17
	$chk_Org_Get_ECPVirtualDirectory.Text = "Get-EcpVirtualDirectory (E14)"
	$chk_Org_Get_ECPVirtualDirectory.UseVisualStyleBackColor = $True
	$bx_ClientAccess_Functions.Controls.Add($chk_Org_Get_ECPVirtualDirectory)
	$chk_Org_Get_OABVirtualDirectory.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_OABVirtualDirectory.Location = $System_Drawing_Point
	$chk_Org_Get_OABVirtualDirectory.Name = "chk_Org_Get_OABVirtualDirectory"
	$chk_Org_Get_OABVirtualDirectory.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_OABVirtualDirectory.TabIndex = 0
	$chk_Org_Get_OABVirtualDirectory.Text = "Get-OabVirtualDirectory"
	$chk_Org_Get_OABVirtualDirectory.UseVisualStyleBackColor = $True
	$bx_ClientAccess_Functions.Controls.Add($chk_Org_Get_OABVirtualDirectory)
	$chk_Org_Get_OutlookAnywhere.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_OutlookAnywhere.Location = $System_Drawing_Point
	$chk_Org_Get_OutlookAnywhere.Name = "chk_Org_Get_OutlookAnywhere"
	$chk_Org_Get_OutlookAnywhere.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_OutlookAnywhere.TabIndex = 2
	$chk_Org_Get_OutlookAnywhere.Text = "Get-OutlookAnywhere"
	$chk_Org_Get_OutlookAnywhere.UseVisualStyleBackColor = $True
	$bx_ClientAccess_Functions.Controls.Add($chk_Org_Get_OutlookAnywhere)
	$chk_Org_Get_OwaMailboxPolicy.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_OwaMailboxPolicy.Location = $System_Drawing_Point
	$chk_Org_Get_OwaMailboxPolicy.Name = "chk_Org_Get_OwaMailboxPolicy"
	$chk_Org_Get_OwaMailboxPolicy.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_OwaMailboxPolicy.TabIndex = 3
	$chk_Org_Get_OwaMailboxPolicy.Text = "Get-OwaMailboxPolicy"
	$chk_Org_Get_OwaMailboxPolicy.UseVisualStyleBackColor = $True
	$bx_ClientAccess_Functions.Controls.Add($chk_Org_Get_OwaMailboxPolicy)
	$chk_Org_Get_OWAVirtualDirectory.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_OWAVirtualDirectory.Location = $System_Drawing_Point
	$chk_Org_Get_OWAVirtualDirectory.Name = "chk_Org_Get_OwaVirtualDirectory"
	$chk_Org_Get_OWAVirtualDirectory.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_OWAVirtualDirectory.TabIndex = 3
	$chk_Org_Get_OWAVirtualDirectory.Text = "Get-OWAVirtualDirectory"
	$chk_Org_Get_OWAVirtualDirectory.UseVisualStyleBackColor = $True
	$bx_ClientAccess_Functions.Controls.Add($chk_Org_Get_OWAVirtualDirectory)
	$chk_Org_Get_PowershellVirtualDirectory.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_2_loc
		$System_Drawing_Point.Y = $Row_2_loc
		$Row_2_loc += 25
	$chk_Org_Get_PowershellVirtualDirectory.Location = $System_Drawing_Point
	$chk_Org_Get_PowershellVirtualDirectory.Name = "chk_Org_Get_PowershellVirtualDirectory"
	$chk_Org_Get_PowershellVirtualDirectory.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_PowershellVirtualDirectory.TabIndex = 4
	$chk_Org_Get_PowershellVirtualDirectory.Text = "Get-PowershellVirtualDir (E14)"
	$chk_Org_Get_PowershellVirtualDirectory.UseVisualStyleBackColor = $True
	$bx_ClientAccess_Functions.Controls.Add($chk_Org_Get_PowershellVirtualDirectory)
	$chk_Org_Get_RPCClientAccess.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_2_loc
		$System_Drawing_Point.Y = $Row_2_loc
		$Row_2_loc += 25
	$chk_Org_Get_RPCClientAccess.Location = $System_Drawing_Point
	$chk_Org_Get_RPCClientAccess.Name = "chk_Org_Get_RPCClientAccess"
	$chk_Org_Get_RPCClientAccess.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_RPCClientAccess.TabIndex = 10
	$chk_Org_Get_RPCClientAccess.Text = "Get-RpcClientAccess (E14)"
	$chk_Org_Get_RPCClientAccess.UseVisualStyleBackColor = $True
	$bx_ClientAccess_Functions.Controls.Add($chk_Org_Get_RPCClientAccess)
	$chk_Org_Get_ThrottlingPolicy.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_2_loc
		$System_Drawing_Point.Y = $Row_2_loc
		$Row_2_loc += 25
	$chk_Org_Get_ThrottlingPolicy.Location = $System_Drawing_Point
	$chk_Org_Get_ThrottlingPolicy.Name = "chk_Org_Get_ThrottlingPolicy"
	$chk_Org_Get_ThrottlingPolicy.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_ThrottlingPolicy.TabIndex = 10
	$chk_Org_Get_ThrottlingPolicy.Text = "Get-ThrottlingPolicy (E14)"
	$chk_Org_Get_ThrottlingPolicy.UseVisualStyleBackColor = $True
	$bx_ClientAccess_Functions.Controls.Add($chk_Org_Get_ThrottlingPolicy)
	$chk_Org_Get_WebServicesVirtualDirectory.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_2_loc
		$System_Drawing_Point.Y = $Row_2_loc
		$Row_2_loc += 25
	$chk_Org_Get_WebServicesVirtualDirectory.Location = $System_Drawing_Point
	$chk_Org_Get_WebServicesVirtualDirectory.Name = "chk_Org_Get_WebServicesVirtualDirectory"
	$chk_Org_Get_WebServicesVirtualDirectory.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_WebServicesVirtualDirectory.TabIndex = 14
	$chk_Org_Get_WebServicesVirtualDirectory.Text = "Get-WebServicesVirtualDirectory"
	$chk_Org_Get_WebServicesVirtualDirectory.UseVisualStyleBackColor = $True
	$bx_ClientAccess_Functions.Controls.Add($chk_Org_Get_WebServicesVirtualDirectory)
}
#EndRegion Step3 Client Access tab

#Region Step3 Global tab
if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true) -or ($NoGUI -eq $true))
{
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = 4
		$System_Drawing_Point.Y = 33
	$tab_Step3_Global.Location = $System_Drawing_Point
	$tab_Step3_Global.Name = "tab_Step3_Global"
	$tab_Step3_Global.Padding = $System_Windows_Forms_Padding_Reusable
		$System_Drawing_Size = New-Object System.Drawing.Size
		$System_Drawing_Size.Height = 488
		$System_Drawing_Size.Width = 542
	$tab_Step3_Global.Size = $System_Drawing_Size
	$tab_Step3_Global.TabIndex = 3
	$tab_Step3_Global.Text = "Global and Database"
	$tab_Step3_Global.UseVisualStyleBackColor = $True
	$tab_Step3_ExOrg_Tier2.Controls.Add($tab_Step3_Global)
	$bx_Global_Functions.Dock = 5
	$bx_Global_Functions.Font = $font_Calibri_10pt_bold
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = 3
		$System_Drawing_Point.Y = 3
	$bx_Global_Functions.Location = $System_Drawing_Point
	$bx_Global_Functions.Name = "bx_Global_Functions"
	$bx_Global_Functions.Size = $System_Drawing_Size_Step3_box
	$bx_Global_Functions.TabIndex = 9
	$bx_Global_Functions.TabStop = $False
	$tab_Step3_Global.Controls.Add($bx_Global_Functions)
	$btn_Step3_Global_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Global_CheckAll.Location = $System_Drawing_Point_Step3_Check
	$btn_Step3_Global_CheckAll.Name = "btn_Step3_Global_CheckAll"
	$btn_Step3_Global_CheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_Global_CheckAll.TabIndex = 28
	$btn_Step3_Global_CheckAll.Text = "Check all on this tab"
	$btn_Step3_Global_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Global_CheckAll.add_Click($handler_btn_Step3_Global_CheckAll_Click)
	$bx_Global_Functions.Controls.Add($btn_Step3_Global_CheckAll)
	$btn_Step3_Global_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Global_UncheckAll.Location = $System_Drawing_Point_Step3_Uncheck
	$btn_Step3_Global_UncheckAll.Name = "btn_Step3_Global_UncheckAll"
	$btn_Step3_Global_UncheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_Global_UncheckAll.TabIndex = 29
	$btn_Step3_Global_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step3_Global_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Global_UncheckAll.add_Click($handler_btn_Step3_Global_UncheckAll_Click)
	$bx_Global_Functions.Controls.Add($btn_Step3_Global_UncheckAll)
	$Col_1_loc = 35
	$Col_2_loc = 290
	$Row_1_loc = 25
	$Row_2_loc = 25
	$chk_Org_Get_AddressBookPolicy.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_AddressBookPolicy.Location = $System_Drawing_Point
	$chk_Org_Get_AddressBookPolicy.Name = "chk_Org_Get_AddressBookPolicy"
	$chk_Org_Get_AddressBookPolicy.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_AddressBookPolicy.TabIndex = 13
	$chk_Org_Get_AddressBookPolicy.Text = "Get-AddressBookPolicy"
	$chk_Org_Get_AddressBookPolicy.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_AddressBookPolicy)
	$chk_Org_Get_AddressList.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_AddressList.Location = $System_Drawing_Point
	$chk_Org_Get_AddressList.Name = "chk_Org_Get_AddressList"
	$chk_Org_Get_AddressList.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_AddressList.TabIndex = 13
	$chk_Org_Get_AddressList.Text = "Get-AddressList"
	$chk_Org_Get_AddressList.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_AddressList)
	$chk_Org_Get_DatabaseAvailabilityGroup.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_DatabaseAvailabilityGroup.Location = $System_Drawing_Point
	$chk_Org_Get_DatabaseAvailabilityGroup.Name = "chk_Org_Get_DatabaseAvailabilityGroup"
	$chk_Org_Get_DatabaseAvailabilityGroup.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_DatabaseAvailabilityGroup.TabIndex = 13
	$chk_Org_Get_DatabaseAvailabilityGroup.Text = "Get-DatabaseAvailabilityGroup (E14)"
	$chk_Org_Get_DatabaseAvailabilityGroup.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_DatabaseAvailabilityGroup)
	$chk_Org_Get_DAGNetwork.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_DAGNetwork.Location = $System_Drawing_Point
	$chk_Org_Get_DAGNetwork.Name = "chk_Org_Get_DAGNetwork"
	$chk_Org_Get_DAGNetwork.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_DAGNetwork.TabIndex = 14
	$chk_Org_Get_DAGNetwork.Text = "Get-DAGNetwork (E14)"
	$chk_Org_Get_DAGNetwork.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_DAGNetwork)
	$chk_Org_Get_EmailAddressPolicy.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_EmailAddressPolicy.Location = $System_Drawing_Point
	$chk_Org_Get_EmailAddressPolicy.Name = "chk_Org_Get_EmailAddressPolicy"
	$chk_Org_Get_EmailAddressPolicy.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_EmailAddressPolicy.TabIndex = 18
	$chk_Org_Get_EmailAddressPolicy.Text = "Get-EmailAddressPolicy"
	$chk_Org_Get_EmailAddressPolicy.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_EmailAddressPolicy)
	$chk_Org_Get_ExchangeCertificate.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_ExchangeCertificate.Location = $System_Drawing_Point
	$chk_Org_Get_ExchangeCertificate.Name = "chk_Org_Get_ExchangeCertificate"
	$chk_Org_Get_ExchangeCertificate.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_ExchangeCertificate.TabIndex = 19
	$chk_Org_Get_ExchangeCertificate.Text = "Get-ExchangeCertificate (E14)"
	$chk_Org_Get_ExchangeCertificate.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_ExchangeCertificate)
	$chk_Org_Get_ExchangeServer.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_ExchangeServer.Location = $System_Drawing_Point
	$chk_Org_Get_ExchangeServer.Name = "chk_Org_Get_ExchangeServer"
	$chk_Org_Get_ExchangeServer.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_ExchangeServer.TabIndex = 20
	$chk_Org_Get_ExchangeServer.Text = "Get-ExchangeServer"
	$chk_Org_Get_ExchangeServer.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_ExchangeServer)
	$chk_Org_Get_MailboxDatabase.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_MailboxDatabase.Location = $System_Drawing_Point
	$chk_Org_Get_MailboxDatabase.Name = "chk_Org_Get_MailboxDatabase"
	$chk_Org_Get_MailboxDatabase.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_MailboxDatabase.TabIndex = 22
	$chk_Org_Get_MailboxDatabase.Text = "Get-MailboxDatabase"
	$chk_Org_Get_MailboxDatabase.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_MailboxDatabase)
	$chk_Org_Get_MailboxDatabaseCopyStatus.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_MailboxDatabaseCopyStatus.Location = $System_Drawing_Point
	$chk_Org_Get_MailboxDatabaseCopyStatus.Name = "chk_Org_Get_MailboxDatabaseCopyStatus"
	$chk_Org_Get_MailboxDatabaseCopyStatus.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_MailboxDatabaseCopyStatus.TabIndex = 23
	$chk_Org_Get_MailboxDatabaseCopyStatus.Text = "Get-MailboxDbCopyStatus (E14)"
	$chk_Org_Get_MailboxDatabaseCopyStatus.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_MailboxDatabaseCopyStatus)
	$chk_Org_Get_MailboxServer.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_MailboxServer.Location = $System_Drawing_Point
	$chk_Org_Get_MailboxServer.Name = "chk_Org_Get_MailboxServer"
	$chk_Org_Get_MailboxServer.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_MailboxServer.TabIndex = 26
	$chk_Org_Get_MailboxServer.Text = "Get-MailboxServer"
	$chk_Org_Get_MailboxServer.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_MailboxServer)
	$chk_Org_Get_OfflineAddressBook.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_OfflineAddressBook.Location = $System_Drawing_Point
	$chk_Org_Get_OfflineAddressBook.Name = "chk_Org_Get_OfflineAddressBook"
	$chk_Org_Get_OfflineAddressBook.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_OfflineAddressBook.TabIndex = 1
	$chk_Org_Get_OfflineAddressBook.Text = "Get-OfflineAddressBook"
	$chk_Org_Get_OfflineAddressBook.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_OfflineAddressBook)
	$chk_Org_Get_OrgConfig.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_OrgConfig.Location = $System_Drawing_Point
	$chk_Org_Get_OrgConfig.Name = "chk_Org_Get_OrgConfig"
	$chk_Org_Get_OrgConfig.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_OrgConfig.TabIndex = 1
	$chk_Org_Get_OrgConfig.Text = "Get-OrganizationConfig"
	$chk_Org_Get_OrgConfig.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_OrgConfig)
	$chk_Org_Get_PublicFolderDatabase.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_2_loc
		$System_Drawing_Point.Y = $Row_2_loc
		$Row_2_loc += 25
	$chk_Org_Get_PublicFolderDatabase.Location = $System_Drawing_Point
	$chk_Org_Get_PublicFolderDatabase.Name = "chk_Org_Get_PublicFolderDatabase"
	$chk_Org_Get_PublicFolderDatabase.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_PublicFolderDatabase.TabIndex = 6
	$chk_Org_Get_PublicFolderDatabase.Text = "Get-PublicFolderDatabase"
	$chk_Org_Get_PublicFolderDatabase.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_PublicFolderDatabase)
	$chk_Org_Get_Rbac.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_2_loc
		$System_Drawing_Point.Y = $Row_2_loc
		$Row_2_loc += 25
	$chk_Org_Get_Rbac.Location = $System_Drawing_Point
	$chk_Org_Get_Rbac.Name = "chk_Org_Get_Rbac"
	$chk_Org_Get_Rbac.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_Rbac.TabIndex = 6
	$chk_Org_Get_Rbac.Text = "Get-Rbac"
	$chk_Org_Get_Rbac.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_Rbac)
	$chk_Org_Get_RetentionPolicy.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_2_loc
		$System_Drawing_Point.Y = $Row_2_loc
		$Row_2_loc += 25
	$chk_Org_Get_RetentionPolicy.Location = $System_Drawing_Point
	$chk_Org_Get_RetentionPolicy.Name = "chk_Org_Get_RetentionPolicy"
	$chk_Org_Get_RetentionPolicy.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_RetentionPolicy.TabIndex = 6
	$chk_Org_Get_RetentionPolicy.Text = "Get-RetentionPolicy"
	$chk_Org_Get_RetentionPolicy.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_RetentionPolicy)
	$chk_Org_Get_RetentionPolicyTag.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_2_loc
		$System_Drawing_Point.Y = $Row_2_loc
		$Row_2_loc += 25
	$chk_Org_Get_RetentionPolicyTag.Location = $System_Drawing_Point
	$chk_Org_Get_RetentionPolicyTag.Name = "chk_Org_Get_RetentionPolicyTag"
	$chk_Org_Get_RetentionPolicyTag.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_RetentionPolicyTag.TabIndex = 6
	$chk_Org_Get_RetentionPolicyTag.Text = "Get-RetentionPolicyTag"
	$chk_Org_Get_RetentionPolicyTag.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_RetentionPolicyTag)
	$chk_Org_Get_StorageGroup.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_2_loc
		$System_Drawing_Point.Y = $Row_2_loc
		$Row_2_loc += 25
	$chk_Org_Get_StorageGroup.Location = $System_Drawing_Point
	$chk_Org_Get_StorageGroup.Name = "chk_Org_Get_StorageGroup"
	$chk_Org_Get_StorageGroup.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_StorageGroup.TabIndex = 11
	$chk_Org_Get_StorageGroup.Text = "Get-StorageGroup (E2k7)"
	$chk_Org_Get_StorageGroup.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_StorageGroup)
}
#EndRegion Step3 Global tab

#Region Step3 Recipient tab
if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true) -or ($NoGUI -eq $true))
{
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = 4
		$System_Drawing_Point.Y = 33
	$tab_Step3_Recipient.Location = $System_Drawing_Point
	$tab_Step3_Recipient.Name = "tab_Step3_Recipient"
	$tab_Step3_Recipient.Padding = $System_Windows_Forms_Padding_Reusable
		$System_Drawing_Size = New-Object System.Drawing.Size
		$System_Drawing_Size.Height = 488
		$System_Drawing_Size.Width = 542
	$tab_Step3_Recipient.Size = $System_Drawing_Size
	$tab_Step3_Recipient.TabIndex = 3
	$tab_Step3_Recipient.Text = "Recipient"
	$tab_Step3_Recipient.UseVisualStyleBackColor = $True
	$tab_Step3_ExOrg_Tier2.Controls.Add($tab_Step3_Recipient)
	$bx_Recipient_Functions.Dock = 5
	$bx_Recipient_Functions.Font = $font_Calibri_10pt_bold
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = 3
		$System_Drawing_Point.Y = 3
	$bx_Recipient_Functions.Location = $System_Drawing_Point
	$bx_Recipient_Functions.Name = "bx_Recipient_Functions"
	$bx_Recipient_Functions.Size = $System_Drawing_Size_Step3_box
	$bx_Recipient_Functions.TabIndex = 9
	$bx_Recipient_Functions.TabStop = $False
	$tab_Step3_Recipient.Controls.Add($bx_Recipient_Functions)
	$btn_Step3_Recipient_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Recipient_CheckAll.Location = $System_Drawing_Point_Step3_Check
	$btn_Step3_Recipient_CheckAll.Name = "btn_Step3_Recipient_CheckAll"
	$btn_Step3_Recipient_CheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_Recipient_CheckAll.TabIndex = 28
	$btn_Step3_Recipient_CheckAll.Text = "Check all on this tab"
	$btn_Step3_Recipient_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Recipient_CheckAll.add_Click($handler_btn_Step3_Recipient_CheckAll_Click)
	$bx_Recipient_Functions.Controls.Add($btn_Step3_Recipient_CheckAll)
	$btn_Step3_Recipient_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Recipient_UncheckAll.Location = $System_Drawing_Point_Step3_Uncheck
	$btn_Step3_Recipient_UncheckAll.Name = "btn_Step3_Recipient_UncheckAll"
	$btn_Step3_Recipient_UncheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_Recipient_UncheckAll.TabIndex = 29
	$btn_Step3_Recipient_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step3_Recipient_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Recipient_UncheckAll.add_Click($handler_btn_Step3_Recipient_UncheckAll_Click)
	$bx_Recipient_Functions.Controls.Add($btn_Step3_Recipient_UncheckAll)
	$Col_1_loc = 35
	$Col_2_loc = 290
	$Row_1_loc = 25
	$Row_2_loc = 25
	$chk_Org_Get_ADPermission.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_ADPermission.Location = $System_Drawing_Point
	$chk_Org_Get_ADPermission.Name = "chk_Org_Get_ADPermission"
	$chk_Org_Get_ADPermission.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_ADPermission.TabIndex = 4
	$chk_Org_Get_ADPermission.Text = "Get-ADPermission"
	$chk_Org_Get_ADPermission.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_ADPermission)
	$chk_Org_Get_CalendarProcessing.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_CalendarProcessing.Location = $System_Drawing_Point
	$chk_Org_Get_CalendarProcessing.Name = "chk_Org_Get_CalendarProcessing"
	$chk_Org_Get_CalendarProcessing.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_CalendarProcessing.TabIndex = 8
	$chk_Org_Get_CalendarProcessing.Text = "Get-CalendarProcessing (E14)"
	$chk_Org_Get_CalendarProcessing.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_CalendarProcessing)
	$chk_Org_Get_CASMailbox.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_CASMailbox.Location = $System_Drawing_Point
	$chk_Org_Get_CASMailbox.Name = "chk_Org_Get_CASMailbox"
	$chk_Org_Get_CASMailbox.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_CASMailbox.TabIndex = 9
	$chk_Org_Get_CASMailbox.Text = "Get-CASMailbox"
	$chk_Org_Get_CASMailbox.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_CASMailbox)
	$chk_Org_Get_DistributionGroup.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_DistributionGroup.Location = $System_Drawing_Point
	$chk_Org_Get_DistributionGroup.Name = "chk_Org_Get_DistributionGroup"
	$chk_Org_Get_DistributionGroup.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_DistributionGroup.TabIndex = 15
	$chk_Org_Get_DistributionGroup.Text = "Get-DistributionGroup"
	$chk_Org_Get_DistributionGroup.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_DistributionGroup)
	$chk_Org_Get_DynamicDistributionGroup.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_DynamicDistributionGroup.Location = $System_Drawing_Point
	$chk_Org_Get_DynamicDistributionGroup.Name = "chk_Org_Get_DynamicDistributionGroup"
	$chk_Org_Get_DynamicDistributionGroup.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_DynamicDistributionGroup.TabIndex = 16
	$chk_Org_Get_DynamicDistributionGroup.Text = "Get-DynamicDistributionGroup"
	$chk_Org_Get_DynamicDistributionGroup.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_DynamicDistributionGroup)
	$chk_Org_Get_Mailbox.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_Mailbox.Location = $System_Drawing_Point
	$chk_Org_Get_Mailbox.Name = "chk_Org_Get_Mailbox"
	$chk_Org_Get_Mailbox.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_Mailbox.TabIndex = 21
	$chk_Org_Get_Mailbox.Text = "Get-Mailbox"
	$chk_Org_Get_Mailbox.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_Mailbox)
	$chk_Org_Get_MailboxFolderStatistics.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_MailboxFolderStatistics.Location = $System_Drawing_Point
	$chk_Org_Get_MailboxFolderStatistics.Name = "chk_Org_Get_MailboxFolderStatistics"
	$chk_Org_Get_MailboxFolderStatistics.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_MailboxFolderStatistics.TabIndex = 24
	$chk_Org_Get_MailboxFolderStatistics.Text = "Get-MailboxFolderStatistics"
	$chk_Org_Get_MailboxFolderStatistics.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_MailboxFolderStatistics)
	$chk_Org_Get_MailboxPermission.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_MailboxPermission.Location = $System_Drawing_Point
	$chk_Org_Get_MailboxPermission.Name = "chk_Org_Get_MailboxPermission"
	$chk_Org_Get_MailboxPermission.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_MailboxPermission.TabIndex = 25
	$chk_Org_Get_MailboxPermission.Text = "Get-MailboxPermission"
	$chk_Org_Get_MailboxPermission.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_MailboxPermission)
	$chk_Org_Get_MailboxStatistics.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_MailboxStatistics.Location = $System_Drawing_Point
	$chk_Org_Get_MailboxStatistics.Name = "chk_Org_Get_MailboxStatistics"
	$chk_Org_Get_MailboxStatistics.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_MailboxStatistics.TabIndex = 27
	$chk_Org_Get_MailboxStatistics.Text = "Get-MailboxStatistics"
	$chk_Org_Get_MailboxStatistics.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_MailboxStatistics)
	$chk_Org_Get_PublicFolder.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_PublicFolder.Location = $System_Drawing_Point
	$chk_Org_Get_PublicFolder.Name = "chk_Org_Get_PublicFolder"
	$chk_Org_Get_PublicFolder.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_PublicFolder.TabIndex = 5
	$chk_Org_Get_PublicFolder.Text = "Get-PublicFolder"
	$chk_Org_Get_PublicFolder.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_PublicFolder)
	$chk_Org_Get_PublicFolderStatistics.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_PublicFolderStatistics.Location = $System_Drawing_Point
	$chk_Org_Get_PublicFolderStatistics.Name = "chk_Org_Get_PublicFolderStatistics"
	$chk_Org_Get_PublicFolderStatistics.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_PublicFolderStatistics.TabIndex = 7
	$chk_Org_Get_PublicFolderStatistics.Text = "Get-PublicFolderStatistics"
	$chk_Org_Get_PublicFolderStatistics.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_PublicFolderStatistics)
	$chk_Org_Get_User.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_User.Location = $System_Drawing_Point
	$chk_Org_Get_User.Name = "$chk_Org_Get_User"
	$chk_Org_Get_User.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_User.TabIndex = 16
	$chk_Org_Get_User.Text = "Get-User"
	$chk_Org_Get_User.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_User)
	$chk_Org_Quota.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Quota.Location = $System_Drawing_Point
	$chk_Org_Quota.Name = "chk_Org_Quota"
	$chk_Org_Quota.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Quota.TabIndex = 15
	$chk_Org_Quota.Text = "Quota"
	$chk_Org_Quota.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Quota)
}
#EndRegion Step3 Recipient tab

#Region Step3 Transport tab
if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true) -or ($NoGUI -eq $true))
{
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = 4
		$System_Drawing_Point.Y = 33
	$tab_Step3_Transport.Location = $System_Drawing_Point
	$tab_Step3_Transport.Name = "tab_Step3_Transport"
	$tab_Step3_Transport.Padding = $System_Windows_Forms_Padding_Reusable
		$System_Drawing_Size = New-Object System.Drawing.Size
		$System_Drawing_Size.Height = 488
		$System_Drawing_Size.Width = 542
	$tab_Step3_Transport.Size = $System_Drawing_Size
	$tab_Step3_Transport.TabIndex = 3
	$tab_Step3_Transport.Text = "Transport"
	$tab_Step3_Transport.UseVisualStyleBackColor = $True
	$tab_Step3_ExOrg_Tier2.Controls.Add($tab_Step3_Transport)
	$bx_Transport_Functions.Dock = 5
	$bx_Transport_Functions.Font = $font_Calibri_10pt_bold
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = 3
		$System_Drawing_Point.Y = 3
	$bx_Transport_Functions.Location = $System_Drawing_Point
	$bx_Transport_Functions.Name = "bx_Transport_Functions"
	$bx_Transport_Functions.Size = $System_Drawing_Size_Step3_box
	$bx_Transport_Functions.TabIndex = 9
	$bx_Transport_Functions.TabStop = $False
	$tab_Step3_Transport.Controls.Add($bx_Transport_Functions)
	$btn_Step3_Transport_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Transport_CheckAll.Location = $System_Drawing_Point_Step3_Check
	$btn_Step3_Transport_CheckAll.Name = "btn_Step3_Transport_CheckAll"
	$btn_Step3_Transport_CheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_Transport_CheckAll.TabIndex = 28
	$btn_Step3_Transport_CheckAll.Text = "Check all on this tab"
	$btn_Step3_Transport_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Transport_CheckAll.add_Click($handler_btn_Step3_Transport_CheckAll_Click)
	$bx_Transport_Functions.Controls.Add($btn_Step3_Transport_CheckAll)
	$btn_Step3_Transport_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Transport_UncheckAll.Location = $System_Drawing_Point_Step3_Uncheck
	$btn_Step3_Transport_UncheckAll.Name = "btn_Step3_Transport_UncheckAll"
	$btn_Step3_Transport_UncheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_Transport_UncheckAll.TabIndex = 29
	$btn_Step3_Transport_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step3_Transport_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Transport_UncheckAll.add_Click($handler_btn_Step3_Transport_UncheckAll_Click)
	$bx_Transport_Functions.Controls.Add($btn_Step3_Transport_UncheckAll)
	$Col_1_loc = 35
	$Col_2_loc = 290
	$Row_1_loc = 25
	$Row_2_loc = 25
	$chk_Org_Get_AcceptedDomain.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_AcceptedDomain.Location = $System_Drawing_Point
	$chk_Org_Get_AcceptedDomain.Name = "chk_Org_Get_AcceptedDomain"
	$chk_Org_Get_AcceptedDomain.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_AcceptedDomain.TabIndex = 0
	$chk_Org_Get_AcceptedDomain.Text = "Get-AcceptedDomain"
	$chk_Org_Get_AcceptedDomain.UseVisualStyleBackColor = $True
	$bx_Transport_Functions.Controls.Add($chk_Org_Get_AcceptedDomain)
	$chk_Org_Get_AdSite.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_AdSite.Location = $System_Drawing_Point
	$chk_Org_Get_AdSite.Name = "chk_Org_Get_AdSite"
	$chk_Org_Get_AdSite.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_AdSite.TabIndex = 5
	$chk_Org_Get_AdSite.Text = "Get-AdSite"
	$chk_Org_Get_AdSite.UseVisualStyleBackColor = $True
	$bx_Transport_Functions.Controls.Add($chk_Org_Get_AdSite)
	$chk_Org_Get_AdSiteLink.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_AdSiteLink.Location = $System_Drawing_Point
	$chk_Org_Get_AdSiteLink.Name = "chk_Org_Get_AdSiteLink"
	$chk_Org_Get_AdSiteLink.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_AdSiteLink.TabIndex = 6
	$chk_Org_Get_AdSiteLink.Text = "Get-AdSiteLink"
	$chk_Org_Get_AdSiteLink.UseVisualStyleBackColor = $True
	$bx_Transport_Functions.Controls.Add($chk_Org_Get_AdSiteLink)
	$chk_Org_Get_ContentFilterConfig.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_ContentFilterConfig.Location = $System_Drawing_Point
	$chk_Org_Get_ContentFilterConfig.Name = "chk_Org_Get_ContentFilterConfig"
	$chk_Org_Get_ContentFilterConfig.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_ContentFilterConfig.TabIndex = 12
	$chk_Org_Get_ContentFilterConfig.Text = "Get-ContentFilterConfig"
	$chk_Org_Get_ContentFilterConfig.UseVisualStyleBackColor = $True
	$bx_Transport_Functions.Controls.Add($chk_Org_Get_ContentFilterConfig)
	$chk_Org_Get_ReceiveConnector.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_ReceiveConnector.Location = $System_Drawing_Point
	$chk_Org_Get_ReceiveConnector.Name = "chk_Org_Get_ReceiveConnector"
	$chk_Org_Get_ReceiveConnector.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_ReceiveConnector.TabIndex = 8
	$chk_Org_Get_ReceiveConnector.Text = "Get-ReceiveConnector"
	$chk_Org_Get_ReceiveConnector.UseVisualStyleBackColor = $True
	$bx_Transport_Functions.Controls.Add($chk_Org_Get_ReceiveConnector)
	$chk_Org_Get_RemoteDomain.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_RemoteDomain.Location = $System_Drawing_Point
	$chk_Org_Get_RemoteDomain.Name = "chk_Org_Get_RemoteDomain"
	$chk_Org_Get_RemoteDomain.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_RemoteDomain.TabIndex = 8
	$chk_Org_Get_RemoteDomain.Text = "Get-RemoteDomain"
	$chk_Org_Get_RemoteDomain.UseVisualStyleBackColor = $True
	$bx_Transport_Functions.Controls.Add($chk_Org_Get_RemoteDomain)
	$chk_Org_Get_RoutingGroupConnector.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_RoutingGroupConnector.Location = $System_Drawing_Point
	$chk_Org_Get_RoutingGroupConnector.Name = "chk_Org_Get_RoutingGroupConnector"
	$chk_Org_Get_RoutingGroupConnector.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_RoutingGroupConnector.TabIndex = 9
	$chk_Org_Get_RoutingGroupConnector.Text = "Get-RoutingGroupConnector"
	$chk_Org_Get_RoutingGroupConnector.UseVisualStyleBackColor = $True
	$bx_Transport_Functions.Controls.Add($chk_Org_Get_RoutingGroupConnector)
	$chk_Org_Get_SendConnector.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_SendConnector.Location = $System_Drawing_Point
	$chk_Org_Get_SendConnector.Name = "chk_Org_Get_SendConnector"
	$chk_Org_Get_SendConnector.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_SendConnector.TabIndex = 11
	$chk_Org_Get_SendConnector.Text = "Get-SendConnector"
	$chk_Org_Get_SendConnector.UseVisualStyleBackColor = $True
	$bx_Transport_Functions.Controls.Add($chk_Org_Get_SendConnector)
	$chk_Org_Get_TransportConfig.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_TransportConfig.Location = $System_Drawing_Point
	$chk_Org_Get_TransportConfig.Name = "chk_Org_Get_TransportConfig"
	$chk_Org_Get_TransportConfig.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_TransportConfig.TabIndex = 12
	$chk_Org_Get_TransportConfig.Text = "Get-TransportConfig"
	$chk_Org_Get_TransportConfig.UseVisualStyleBackColor = $True
	$bx_Transport_Functions.Controls.Add($chk_Org_Get_TransportConfig)
	$chk_Org_Get_TransportRule.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_TransportRule.Location = $System_Drawing_Point
	$chk_Org_Get_TransportRule.Name = "chk_Org_Get_TransportRule"
	$chk_Org_Get_TransportRule.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_TransportRule.TabIndex = 12
	$chk_Org_Get_TransportRule.Text = "Get-TransportRule"
	$chk_Org_Get_TransportRule.UseVisualStyleBackColor = $True
	$bx_Transport_Functions.Controls.Add($chk_Org_Get_TransportRule)
	$chk_Org_Get_TransportServer.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_TransportServer.Location = $System_Drawing_Point
	$chk_Org_Get_TransportServer.Name = "chk_Org_Get_TransportServer"
	$chk_Org_Get_TransportServer.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_TransportServer.TabIndex = 13
	$chk_Org_Get_TransportServer.Text = "Get-TransportServer"
	$chk_Org_Get_TransportServer.UseVisualStyleBackColor = $True
	$bx_Transport_Functions.Controls.Add($chk_Org_Get_TransportServer)
}
#EndRegion Step3 Transport tab

#Region Step3 UM tab
if ((($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true) -or ($NoGUI -eq $true)) -and ($UM -eq $true))
{
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = 4
		$System_Drawing_Point.Y = 33
	$tab_Step3_UM.Location = $System_Drawing_Point
	$tab_Step3_UM.Name = "tab_Step3_Misc"
	$tab_Step3_UM.Padding = $System_Windows_Forms_Padding_Reusable
		$System_Drawing_Size = New-Object System.Drawing.Size
		$System_Drawing_Size.Height = 488
		$System_Drawing_Size.Width = 300 #542
	$tab_Step3_UM.Size = $System_Drawing_Size
	$tab_Step3_UM.TabIndex = 4
	$tab_Step3_UM.Text = "    UM"
	$tab_Step3_UM.UseVisualStyleBackColor = $True
	$tab_Step3_ExOrg_Tier2.Controls.Add($tab_Step3_UM)
	$bx_UM_Functions.Dock = 5
	$bx_UM_Functions.Font = $font_Calibri_10pt_bold
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = 3
		$System_Drawing_Point.Y = 3
	$bx_UM_Functions.Location = $System_Drawing_Point
	$bx_UM_Functions.Name = "bx_Misc_Functions"
	$bx_UM_Functions.Size = $System_Drawing_Size_Step3_box
	$bx_UM_Functions.TabIndex = 9
	$bx_UM_Functions.TabStop = $False
	$tab_Step3_UM.Controls.Add($bx_UM_Functions)
	$btn_Step3_UM_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_UM_CheckAll.Location = $System_Drawing_Point_Step3_Check
	$btn_Step3_UM_CheckAll.Name = "btn_Step3_Misc_CheckAll"
	$btn_Step3_UM_CheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_UM_CheckAll.TabIndex = 28
	$btn_Step3_UM_CheckAll.Text = "Check all on this tab"
	$btn_Step3_UM_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_UM_CheckAll.add_Click($handler_btn_Step3_UM_CheckAll_Click)
	$bx_UM_Functions.Controls.Add($btn_Step3_UM_CheckAll)
	$btn_Step3_UM_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_UM_UncheckAll.Location = $System_Drawing_Point_Step3_Uncheck
	$btn_Step3_UM_UncheckAll.Name = "btn_Step3_Misc_UncheckAll"
	$btn_Step3_UM_UncheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_UM_UncheckAll.TabIndex = 29
	$btn_Step3_UM_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step3_UM_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_UM_UncheckAll.add_Click($handler_btn_Step3_UM_UncheckAll_Click)
	$bx_UM_Functions.Controls.Add($btn_Step3_UM_UncheckAll)
	$Col_1_loc = 35
	$Col_2_loc = 290
	$Row_1_loc = 25
	$Row_2_loc = 25
	$chk_Org_Get_UmAutoAttendant.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_UmAutoAttendant.Location = $System_Drawing_Point
	$chk_Org_Get_UmAutoAttendant.Name = "chk_Org_Get_UmAutoAttendant"
	$chk_Org_Get_UmAutoAttendant.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_UmAutoAttendant.TabIndex = 0
	$chk_Org_Get_UmAutoAttendant.Text = "Get-UmAutoAttendant"
	$chk_Org_Get_UmAutoAttendant.UseVisualStyleBackColor = $True
	$bx_UM_Functions.Controls.Add($chk_Org_Get_UmAutoAttendant)
	$chk_Org_Get_UmDialPlan.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_UmDialPlan.Location = $System_Drawing_Point
	$chk_Org_Get_UmDialPlan.Name = "chk_Org_Get_UmDialPlan"
	$chk_Org_Get_UmDialPlan.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_UmDialPlan.TabIndex = 0
	$chk_Org_Get_UmDialPlan.Text = "Get-UmDialPlan"
	$chk_Org_Get_UmDialPlan.UseVisualStyleBackColor = $True
	$bx_UM_Functions.Controls.Add($chk_Org_Get_UmDialPlan)
	$chk_Org_Get_UmIpGateway.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_UmIpGateway.Location = $System_Drawing_Point
	$chk_Org_Get_UmIpGateway.Name = "chk_Org_Get_UmIpGateway"
	$chk_Org_Get_UmIpGateway.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_UmIpGateway.TabIndex = 0
	$chk_Org_Get_UmIpGateway.Text = "Get-UmIpGateway"
	$chk_Org_Get_UmIpGateway.UseVisualStyleBackColor = $True
	$bx_UM_Functions.Controls.Add($chk_Org_Get_UmIpGateway)
	$chk_Org_Get_UmMailbox.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_UmMailbox.Location = $System_Drawing_Point
	$chk_Org_Get_UmMailbox.Name = "chk_Org_Get_UmMailbox"
	$chk_Org_Get_UmMailbox.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_UmMailbox.TabIndex = 0
	$chk_Org_Get_UmMailbox.Text = "Get-UmMailbox"
	$chk_Org_Get_UmMailbox.UseVisualStyleBackColor = $True
	$bx_UM_Functions.Controls.Add($chk_Org_Get_UmMailbox)
#	$chk_Org_Get_UmMailboxConfiguration.Font = $font_Calibri_10pt_normal
#		$System_Drawing_Point = New-Object System.Drawing.Point
#		$System_Drawing_Point.X = $Col_1_loc
#		$System_Drawing_Point.Y = $Row_1_loc
#		$Row_1_loc += 25
#	$chk_Org_Get_UmMailboxConfiguration.Location = $System_Drawing_Point
#	$chk_Org_Get_UmMailboxConfiguration.Name = "chk_Org_Get_UmMailboxConfiguration"
#	$chk_Org_Get_UmMailboxConfiguration.Size = $System_Drawing_Size_Reusable_chk
#	$chk_Org_Get_UmMailboxConfiguration.TabIndex = 0
#	$chk_Org_Get_UmMailboxConfiguration.Text = "Get-UmMailboxConfiguration"
#	$chk_Org_Get_UmMailboxConfiguration.UseVisualStyleBackColor = $True
#	$bx_UM_Functions.Controls.Add($chk_Org_Get_UmMailboxConfiguration)

#	$chk_Org_Get_UmMailboxPin.Font = $font_Calibri_10pt_normal
#		$System_Drawing_Point = New-Object System.Drawing.Point
#		$System_Drawing_Point.X = $Col_1_loc
#		$System_Drawing_Point.Y = $Row_1_loc
#		$Row_1_loc += 25
#	$chk_Org_Get_UmMailboxPin.Location = $System_Drawing_Point
#	$chk_Org_Get_UmMailboxPin.Name = "chk_Org_Get_UmMailboxPin"
#	$chk_Org_Get_UmMailboxPin.Size = $System_Drawing_Size_Reusable_chk
#	$chk_Org_Get_UmMailboxPin.TabIndex = 0
#	$chk_Org_Get_UmMailboxPin.Text = "Get-UmMailboxPin"
#	$chk_Org_Get_UmMailboxPin.UseVisualStyleBackColor = $True
#	$bx_UM_Functions.Controls.Add($chk_Org_Get_UmMailboxPin)
	$chk_Org_Get_UmMailboxPolicy.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_UmMailboxPolicy.Location = $System_Drawing_Point
	$chk_Org_Get_UmMailboxPolicy.Name = "chk_Org_Get_UmMailboxPolicy"
	$chk_Org_Get_UmMailboxPolicy.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_UmMailboxPolicy.TabIndex = 0
	$chk_Org_Get_UmMailboxPolicy.Text = "Get-UmMailboxPolicy"
	$chk_Org_Get_UmMailboxPolicy.UseVisualStyleBackColor = $True
	$bx_UM_Functions.Controls.Add($chk_Org_Get_UmMailboxPolicy)
	$chk_Org_Get_UmServer.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_UmServer.Location = $System_Drawing_Point
	$chk_Org_Get_UmServer.Name = "chk_Org_Get_UmServer"
	$chk_Org_Get_UmServer.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_UmServer.TabIndex = 0
	$chk_Org_Get_UmServer.Text = "Get-UmServer"
	$chk_Org_Get_UmServer.UseVisualStyleBackColor = $True
	$bx_UM_Functions.Controls.Add($chk_Org_Get_UmServer)
}
#EndRegion Step3 UM tab

#Region Step3 Misc tab
if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true) -or ($NoGUI -eq $true))
{
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = 4
		$System_Drawing_Point.Y = 33
	$tab_Step3_Misc.Location = $System_Drawing_Point
	$tab_Step3_Misc.Name = "tab_Step3_Misc"
	$tab_Step3_Misc.Padding = $System_Windows_Forms_Padding_Reusable
		$System_Drawing_Size = New-Object System.Drawing.Size
		$System_Drawing_Size.Height = 488
		$System_Drawing_Size.Width = 542
	$tab_Step3_Misc.Size = $System_Drawing_Size
	$tab_Step3_Misc.TabIndex = 4
	$tab_Step3_Misc.Text = "Misc"
	$tab_Step3_Misc.UseVisualStyleBackColor = $True
	$tab_Step3_ExOrg_Tier2.Controls.Add($tab_Step3_Misc)
	$bx_Misc_Functions.Dock = 5
	$bx_Misc_Functions.Font = $font_Calibri_10pt_bold
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = 3
		$System_Drawing_Point.Y = 3
	$bx_Misc_Functions.Location = $System_Drawing_Point
	$bx_Misc_Functions.Name = "bx_Misc_Functions"
	$bx_Misc_Functions.Size = $System_Drawing_Size_Step3_box
	$bx_Misc_Functions.TabIndex = 9
	$bx_Misc_Functions.TabStop = $False
	$tab_Step3_Misc.Controls.Add($bx_Misc_Functions)
	$btn_Step3_Misc_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Misc_CheckAll.Location = $System_Drawing_Point_Step3_Check
	$btn_Step3_Misc_CheckAll.Name = "btn_Step3_Misc_CheckAll"
	$btn_Step3_Misc_CheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_Misc_CheckAll.TabIndex = 28
	$btn_Step3_Misc_CheckAll.Text = "Check all on this tab"
	$btn_Step3_Misc_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Misc_CheckAll.add_Click($handler_btn_Step3_Misc_CheckAll_Click)
	$bx_Misc_Functions.Controls.Add($btn_Step3_Misc_CheckAll)

	$btn_Step3_Misc_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Misc_UncheckAll.Location = $System_Drawing_Point_Step3_Uncheck
	$btn_Step3_Misc_UncheckAll.Name = "btn_Step3_Misc_UncheckAll"
	$btn_Step3_Misc_UncheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_Misc_UncheckAll.TabIndex = 29
	$btn_Step3_Misc_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step3_Misc_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Misc_UncheckAll.add_Click($handler_btn_Step3_Misc_UncheckAll_Click)
	$bx_Misc_Functions.Controls.Add($btn_Step3_Misc_UncheckAll)
	$Col_1_loc = 35
	$Col_2_loc = 290
	$Row_1_loc = 25
	$Row_2_loc = 25
	$chk_Org_Get_AdminGroups.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_AdminGroups.Location = $System_Drawing_Point
	$chk_Org_Get_AdminGroups.Name = "chk_Org_Get_AdminGroups"
	$chk_Org_Get_AdminGroups.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_AdminGroups.TabIndex = 0
	$chk_Org_Get_AdminGroups.Text = "Get memberships of admin groups"
	$chk_Org_Get_AdminGroups.UseVisualStyleBackColor = $True
	$bx_Misc_Functions.Controls.Add($chk_Org_Get_AdminGroups)
	$chk_Org_Get_Fsmo.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_Fsmo.Location = $System_Drawing_Point
	$chk_Org_Get_Fsmo.Name = "chk_Org_Get_Fsmo"
	$chk_Org_Get_Fsmo.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_Fsmo.TabIndex = 0
	$chk_Org_Get_Fsmo.Text = "Get FSMO role holders for domains"
	$chk_Org_Get_Fsmo.UseVisualStyleBackColor = $True
	$bx_Misc_Functions.Controls.Add($chk_Org_Get_Fsmo)
	$chk_Org_Get_ExchangeServerBuilds.Font = $font_Calibri_10pt_normal
		$System_Drawing_Point = New-Object System.Drawing.Point
		$System_Drawing_Point.X = $Col_1_loc
		$System_Drawing_Point.Y = $Row_1_loc
		$Row_1_loc += 25
	$chk_Org_Get_ExchangeServerBuilds.Location = $System_Drawing_Point
	$chk_Org_Get_ExchangeServerBuilds.Name = "chk_Org_Get_ExchangeServerBuilds"
	$chk_Org_Get_ExchangeServerBuilds.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_ExchangeServerBuilds.TabIndex = 0
	$chk_Org_Get_ExchangeServerBuilds.Text = "Get Exchange Server build numbers"
	$chk_Org_Get_ExchangeServerBuilds.UseVisualStyleBackColor = $True
	$bx_Misc_Functions.Controls.Add($chk_Org_Get_ExchangeServerBuilds)
}
#EndRegion Step3 Misc tab

#EndRegion "Step3 - Tests"

#Region "Step4 - Reporting"
$tab_Step4.BackColor = [System.Drawing.Color]::FromArgb(0,255,255,255)
$tab_Step4.Font = $font_Calibri_8pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 36
$tab_Step4.Location = $System_Drawing_Point
$tab_Step4.Name = "tab_Step4"
$tab_Step4.Padding = $System_Windows_Forms_Padding_Reusable
$tab_Step4.TabIndex = 3
$tab_Step4.Text = "  Reporting  "
$tab_Step4.Size = $System_Drawing_Size_tab_1
$tab_Master.Controls.Add($tab_Step4)
$btn_Step4_Assemble.Font = $font_Calibri_14pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 20
	$System_Drawing_Point.Y = 15
$btn_Step4_Assemble.Location = $System_Drawing_Point
$btn_Step4_Assemble.Name = "btn_Step4_Assemble"
$btn_Step4_Assemble.Size = $System_Drawing_Size_buttons
$btn_Step4_Assemble.TabIndex = 10
$btn_Step4_Assemble.Text = "Execute"
$btn_Step4_Assemble.UseVisualStyleBackColor = $True
$btn_Step4_Assemble.add_Click($handler_btn_Step4_Assemble_Click)
$tab_Step4.Controls.Add($btn_Step4_Assemble)
$lbl_Step4_Assemble.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 138
	$System_Drawing_Point.Y = 15
$lbl_Step4_Assemble.Location = $System_Drawing_Point
$lbl_Step4_Assemble.Name = "lbl_Step4"
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 38
	$System_Drawing_Size.Width = 510
$lbl_Step4_Assemble.Size = $System_Drawing_Size
$lbl_Step4_Assemble.TabIndex = 11
$lbl_Step4_Assemble.Text = "If Office 2003 or later is installed, the Execute button can be used to assemble `nthe output from Tests into reports."
$lbl_Step4_Assemble.TextAlign = 16
$tab_Step4.Controls.Add($lbl_Step4_Assemble)
$bx_Step4_Functions.Font = $font_Calibri_10pt_bold
	$System_Drawing_Point_bx_Step4 = New-Object System.Drawing.Point
	$System_Drawing_Point_bx_Step4.X = 27	# 96-69
	$System_Drawing_Point_bx_Step4.Y = 91
$bx_Step4_Functions.Location = $System_Drawing_Point_bx_Step4
$bx_Step4_Functions.Name = "bx_Step4_Functions"
	$System_Drawing_Size_bx_Step4 = New-Object System.Drawing.Size
	$System_Drawing_Size_bx_Step4.Height = 487 #482 to short
	$System_Drawing_Size_bx_Step4.Width = 536
$bx_Step4_Functions.Size = $System_Drawing_Size_bx_Step4
$bx_Step4_Functions.TabIndex = 0
$bx_Step4_Functions.TabStop = $False
$bx_Step4_Functions.Text = "Report Generation Functions"
$tab_Step4.Controls.Add($bx_Step4_Functions)
$chk_Step4_DC_Report.Checked = $True
$chk_Step4_DC_Report.CheckState = 1
$chk_Step4_DC_Report.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 50
	$System_Drawing_Point.Y = 25
$chk_Step4_DC_Report.Location = $System_Drawing_Point
$chk_Step4_DC_Report.Name = "chk_Step4_DC_Report"
$chk_Step4_DC_Report.Size = $System_Drawing_Size_Reusable_chk_long
$chk_Step4_DC_Report.TabIndex = 0
$chk_Step4_DC_Report.Text = "Generate Excel for Domain Controllers"
$chk_Step4_DC_Report.UseVisualStyleBackColor = $True
$bx_Step4_Functions.Controls.Add($chk_Step4_DC_Report)
$chk_Step4_Ex_Report.Checked = $True
$chk_Step4_Ex_Report.CheckState = 1
$chk_Step4_Ex_Report.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 50
	$System_Drawing_Point.Y = 50
$chk_Step4_Ex_Report.Location = $System_Drawing_Point
$chk_Step4_Ex_Report.Name = "chk_Step4_Ex_Report"
$chk_Step4_Ex_Report.Size = $System_Drawing_Size_Reusable_chk_long
$chk_Step4_Ex_Report.TabIndex = 1
$chk_Step4_Ex_Report.Text = "Generate Excel for Exchange servers"
$chk_Step4_Ex_Report.UseVisualStyleBackColor = $True
$bx_Step4_Functions.Controls.Add($chk_Step4_Ex_Report)
$chk_Step4_ExOrg_Report.Checked = $True
$chk_Step4_ExOrg_Report.CheckState = 1
$chk_Step4_ExOrg_Report.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 50
	$System_Drawing_Point.Y = 75
$chk_Step4_ExOrg_Report.Location = $System_Drawing_Point
$chk_Step4_ExOrg_Report.Name = "chk_Step4_ExOrg_Report"
$chk_Step4_ExOrg_Report.Size = $System_Drawing_Size_Reusable_chk_long
$chk_Step4_ExOrg_Report.TabIndex = 2
$chk_Step4_ExOrg_Report.Text = "Generate Excel for Exchange Organization"
$chk_Step4_ExOrg_Report.UseVisualStyleBackColor = $True
$bx_Step4_Functions.Controls.Add($chk_Step4_ExOrg_Report)
$chk_Step4_Exchange_Environment_Doc.Checked = $True
$chk_Step4_Exchange_Environment_Doc.CheckState = 1
$chk_Step4_Exchange_Environment_Doc.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 50
	$System_Drawing_Point.Y = 100
$chk_Step4_Exchange_Environment_Doc.Location = $System_Drawing_Point
$chk_Step4_Exchange_Environment_Doc.Name = "chk_Step4_Exchange_Environment_Doc"
$chk_Step4_Exchange_Environment_Doc.Size = $System_Drawing_Size_Reusable_chk_long
$chk_Step4_Exchange_Environment_Doc.TabIndex = 2
$chk_Step4_Exchange_Environment_Doc.Text = "Generate Word for Exchange Documention"
$chk_Step4_Exchange_Environment_Doc.UseVisualStyleBackColor = $True
$bx_Step4_Functions.Controls.Add($chk_Step4_Exchange_Environment_Doc)
$Status_Step4.Font = $font_Calibri_10pt_normal
$Status_Step4.Location = $System_Drawing_Point_Status
$Status_Step4.Name = "Status_Step4"
$Status_Step4.Size = $System_Drawing_Size_Status
$Status_Step4.TabIndex = 12
$Status_Step4.Text = "Step 4 Status"
$tab_Step4.Controls.Add($Status_Step4)
#EndRegion "Step4 - Reporting"

<#
#Region "Step5 - Having Trouble?"
$tab_Step5.BackColor = [System.Drawing.Color]::FromArgb(0,255,255,255)
$tab_Step5.Font = $font_Calibri_8pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 36
$tab_Step5.Location = $System_Drawing_Point
$tab_Step5.Name = "tab_Step5"
$tab_Step5.Padding = $System_Windows_Forms_Padding_Reusable
$tab_Step5.TabIndex = 3
$tab_Step5.Text = "  Having Trouble?  "
$tab_Step5.Size = $System_Drawing_Size_tab_2
$tab_Step5.visible = $False
$tab_Master.Controls.Add($tab_Step5)

$bx_Step5_Functions.Font = $font_Calibri_10pt_bold
	$System_Drawing_Point_bx_Step5 = New-Object System.Drawing.Point
	$System_Drawing_Point_bx_Step5.X = 27	# 96-69
	$System_Drawing_Point_bx_Step5.Y = 91
$bx_Step5_Functions.Location = $System_Drawing_Point_bx_Step5
$bx_Step5_Functions.Name = "bx_Step5_Functions"
	$System_Drawing_Size_bx_Step5 = New-Object System.Drawing.Size
	$System_Drawing_Size_bx_Step5.Height = 487 #482 to short
	$System_Drawing_Size_bx_Step5.Width = 536
$bx_Step5_Functions.Size = $System_Drawing_Size_bx_Step5
$bx_Step5_Functions.TabIndex = 0
$bx_Step5_Functions.TabStop = $False
$bx_Step5_Functions.Text = "If you're having trouble collecting data..."
$tab_Step5.Controls.Add($bx_Step5_Functions)

$Status_Step5.Font = $font_Calibri_10pt_normal
$Status_Step5.Location = $System_Drawing_Point_Status
$Status_Step5.Name = "Status_Step5"
$Status_Step5.Size = $System_Drawing_Size_Status
$Status_Step5.TabIndex = 12
$Status_Step5.Text = "Step 5 Status"
$tab_Step5.Controls.Add($Status_Step5)

#EndRegion "Step5 - Having Trouble?"
#>

#Region Set Tests Checkbox States
if (($INI_Server -ne "") -or ($INI_Cluster -ne "") -or ($INI_ExOrg -ne ""))
{
	# Code to parse INI
	write-host "Importing INI settings"

	# Server INI
	if (($ini_Server -ne "") -and ((Test-Path $ini_Server) -eq $true))
	{
		write-host "File specified using the -INI_Server switch" -ForegroundColor Green
		& ".\ExDC_Scripts\Core_Parse_Ini_File.ps1" -IniFile $INI_Server
	}
	elseif (($ini_Server -ne "") -and ((Test-Path $ini_Server) -eq $false))
	{
		write-host "File specified using the -INI_Server switch was not found" -ForegroundColor Red
	}

	# Cluster INI
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2003orEarlier -eq $true) -or ($NoEMS -eq $true) -or ($NoGUI -eq $true))
	{
		if (($ini_Cluster -ne "") -and ((Test-Path $ini_Cluster) -eq $true))
		{
			write-host "File specified using the -INI_Cluster switch" -ForegroundColor Green
			& ".\ExDC_Scripts\Core_Parse_Ini_File.ps1" -IniFile $INI_Cluster
		}
		elseif (($ini_Cluster -ne "") -and ((Test-Path $ini_Cluster) -eq $false))
		{
			write-host "File specified using the -INI_Cluster switch was not found" -ForegroundColor Red
		}
	}

	# ExOrg INI
	write-host $ini_ExOrg
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true) -or ($NoGUI -eq $true))
	{
		if (($ini_ExOrg -ne "") -and ((Test-Path $ini_ExOrg) -eq $true))
		{
			write-host "File specified using the -INI_ExOrg switch" -ForegroundColor Green
			& ".\ExDC_Scripts\Core_Parse_Ini_File.ps1" -IniFile $INI_ExOrg
		}
		elseif (($ini_ExOrg -ne "") -and ((Test-Path $ini_ExOrg) -eq $false))
		{
			write-host "File specified using the -INI_ExOrg switch was not found" -ForegroundColor Red
		}
	}
}
else
{

	# Domain Controllers
	Set-AllFunctionsDc -Check $true

	# Exchange Servers
	Set-AllFunctionsEx -Check $true

	# Cluster Nodes
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2003orEarlier -eq $true) -or ($NoEMS -eq $true))
	{
		Set-AllFunctionsCluster -Check $true
	}

	# ExOrg Functions
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true))
	{
		Set-AllFunctionsClientAccess -Check $true
		Set-AllFunctionsGlobal -Check $true
		Set-AllFunctionsRecipient -Check $true
		Set-AllFunctionsTransport -Check $true
		Set-AllFunctionsMisc -Check $true
	}
	if ((($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true)) -and ($UM -eq $true))
	{
		Set-AllFunctionsUm -Check $true
	}
}

#EndRegion Set Checkbox States

#endregion *** Build Form ***

#Save the initial state of the form
	$InitialFormWindowState = $form1.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$form1.add_Load($OnLoadForm_StateCorrection)
If ($NoGUI -ne $true)
{
	#Show the Form
	$form1.ShowDialog()| Out-Null
}
Start-NoGUI
}
#End Function

##############################################
# New-ExDCForm should not be above this line #
##############################################

#region *** Custom Functions ***

Trap {
$ErrorText = "ExDC " + "`n" + $server + "`n"
$ErrorText += $_

$ErrorLog = New-Object System.Diagnostics.EventLog('Application')
$ErrorLog.MachineName = "."
$ErrorLog.Source = "ExDC"
Try{$ErrorLog.WriteEntry($ErrorText,"Error", 100)} catch{}
}

Function Watch-ExDCKnownErrors()
{
	Trap [System.InvalidOperationException]{
		#write-host -Fore Red -back White $_.Exception.Message
		#write-host -Fore Red -back White $_.Exception.FullyQualifiedErrorId
		Continue
	}
	Trap [system.Management.Automation.ErrorRecord] {
		#write-host -Fore Red -back White $_.Exception.Message
		#write-host -Fore Red -back White $_.Exception.FullyQualifiedErrorId
		Continue
	}
	Trap [System.Management.AutomationRuntimeException] {
		write-host -Fore Red -back White $_.Exception.Message
		write-host -Fore Red -back White $_.Exception.FullyQualifiedErrorId
		Silently Continue
	}
	Trap [System.Management.Automation.MethodInvocationException] {
		write-host -Fore Red -back White $_.Exception.Message
		write-host -Fore Red -back White $_.Exception.FullyQualifiedErrorId
		Continue
	}
}

Function Start-Execute()
{
	try
	{
		Start-Transcript -Path (".\ExDC_Step3_Transcript_" + $append + ".txt")
	}
	catch [System.Management.Automation.CmdletInvocationException]
	{
		write-host "Transcription already started" -ForegroundColor red
		write-host "Restarting transcription" -ForegroundColor red
		Stop-Transcript
		Start-Transcript -Path (".\ExDC_Step3_Transcript_" + $append + ".txt")
	}
	$btn_Step3_Execute.enabled = $false
	$status_Step3.Text = "Step 3 Status: Running"
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "ExDC"
	$EventLogText = "Starting ExDC Step 3`nDomain Controllers: $intDCTotal`nExchange Servers: $intExTotal`nMailboxes: $intMailboxTotal"
	try{$EventLog.WriteEntry($EventLogText,"Information", 30)} catch{}
	#send the form to the back to expose the Powershell window when starting Step 3
	$form1.WindowState = "minimized"
	write-host "ExDC Form minimized." -ForegroundColor Green

	#Region Executing Domain Controllers Tests
	write-host "Starting Domain Controllers..." -ForegroundColor Green
	if (Get-DCBoxStatus = $true)
	{
		[array]$array_DC_Checked = $null
		foreach ($item in $clb_Step1_DC_List.checkeditems)
		{
			$array_DC_Checked = $array_DC_Checked + $item.tostring()
		}
		if ($null -ne $array_DC_Checked)
		{
			foreach ($server in $array_DC_Checked)
			{
				$ping_reply = $null
				$ping = New-Object system.Net.NetworkInformation.Ping
				try
				{
					$ping_reply = $ping.send($server)
				}
				catch [system.Net.NetworkInformation.PingException]
				{
					write-host "Exception occured during Ping request to server: $server" -ForegroundColor Red
				}
				if (($ping_reply.status) -eq "Success") #This should check connections
#				if ((Test-Connection $server -count 2 -quiet) -eq $true) #This should check connections
				{
				    If ($chk_DC_Win32_Bios.checked -eq $true)
						{Start-ExDCJob -server $server -job "Win32_BIOS" -JobType 0 -Location $location -JobScriptName "dc_w32_bios.ps1" -i $null -PSSession $null}
				    If ($chk_DC_Win32_ComputerSystem.checked -eq $true)
						{Start-ExDCJob -server $server -job "Win32_ComputerSystem" -JobType 0 -Location $location -JobScriptName "dc_w32_cs.ps1" -i $null -PSSession $null}
					If ($chk_DC_Win32_LogicalDisk.checked -eq $true)
						{Start-ExDCJob -server $server -job "Win32_LogicalDisk" -JobType 0 -Location $location -JobScriptName "dc_w32_ld.ps1" -i $null -PSSession $null}
					If ($chk_DC_Win32_NetworkAdapter.checked -eq $true)
						{Start-ExDCJob -server $server -job "Win32_NetworkAdapter" -JobType 0 -Location $location -JobScriptName "dc_w32_na.ps1" -i $null -PSSession $null}
					If ($chk_DC_Win32_NetworkAdapterConfig.checked -eq $true)
						{Start-ExDCJob -server $server -job "Win32_NetworkAdapterConfiguration" -JobType 0 -Location $location -JobScriptName "dc_w32_nac.ps1" -i $null -PSSession $null}
					If ($chk_DC_Win32_OperatingSystem.checked -eq $true)
						{Start-ExDCJob -server $server -job "Win32_OperatingSystem" -JobType 0 -Location $location -JobScriptName "dc_w32_os.ps1" -i $null -PSSession $null}
				    If ($chk_DC_Win32_PageFileUsage.checked -eq $true)
						{Start-ExDCJob -server $server -job "Win32_PageFileUsage" -JobType 0 -Location $location -JobScriptName "dc_w32_pfu.ps1" -i $null -PSSession $null}
				    If ($chk_DC_Win32_PhysicalMemory.checked -eq $true)
						{Start-ExDCJob -server $server -job "Win32_PhysicalMemory" -JobType 0 -Location $location -JobScriptName "dc_w32_pm.ps1" -i $null -PSSession $null}
					If ($chk_DC_Win32_Processor.checked -eq $true)
						{Start-ExDCJob -server $server -job "Win32_Processor" -JobType 0 -Location $location -JobScriptName "dc_w32_proc.ps1" -i $null -PSSession $null}
					If ($chk_DC_Registry_AD.checked -eq $true)
						{Start-ExDCJob -server $server -job "Registry - AD" -JobType 0 -Location $location -JobScriptName "dc_reg_AD.ps1" -i $null -PSSession $null}
					If ($chk_DC_Registry_OS.checked -eq $true)
						{Start-ExDCJob -server $server -job "Registry - OS" -JobType 0 -Location $location -JobScriptName "dc_reg_OS.ps1" -i $null -PSSession $null}
					If ($chk_DC_Registry_Software.checked -eq $true)
						{Start-ExDCJob -server $server -job "Registry - Software" -JobType 0 -Location $location -JobScriptName "dc_reg_Software.ps1" -i $null -PSSession $null}
					If ($chk_DC_MicrosoftDNS_Zone.checked -eq $true)
						{Start-ExDCJob -server $server -job "MicrosoftDNS_Zone" -JobType 0 -Location $location -JobScriptName "dc_dns.ps1" -i $null -PSSession $null}
				    If ($chk_DC_MSAD_DomainController.checked -eq $true)
						{Start-ExDCJob -server $server -job "MSAD_DomainController" -JobType 0 -Location $location -JobScriptName "dc_MSAD_DomainController.ps1" -i $null -PSSession $null}
				    If ($chk_DC_MSAD_ReplNeighbor.checked -eq $true)
						{Start-ExDCJob -server $server -job "MSAD_ReplNeighbor" -JobType 0 -Location $location -JobScriptName "dc_MSAD_ReplNeighbor.ps1" -i $null -PSSession $null}
				}
				else
				{
					$FailedPingOutput = ".\FailedPing_" + $append + ".txt"
					if ((Test-Path $FailedPingOutput) -eq $false)
			        {
				      new-item $FailedPingOutput -type file -Force
			        }
			        "* * * * * * * * * * * * * * * * * * * *"  | Out-File $FailedPingOutput -Force -Append
					write-host "---- Server $server failed Test-Connection for Domain Controller WMI functions" -ForegroundColor Red
					"Server: " + $server  | Out-File $FailedPingOutput -Force -Append
					"Failure: Ping reply to $server was not successful`n`n" | Out-File $FailedPingOutput -Force -Append
#					"Failure: Test-Connection $server -count 2 -quiet`n`n" | Out-File $FailedPingOutput -Force -Append
					$ErrorText = "ExDC " + "`n" + $server + "`n"
					$ErrorText += "Test-Connection failed"
					$ErrorLog = New-Object System.Diagnostics.EventLog('Application')
					$ErrorLog.MachineName = "."
					$ErrorLog.Source = "ExDC"
					Try{$ErrorLog.WriteEntry($ErrorText,"Error", 400)} catch{}
				}
			}
		}
		else
		{
			write-host "No Domain Controllers in the list are checked"
		}
	}
	else
	{
		write-host "---- No Domain Controller Functions selected"
	}
	#EndRegion Executing Domain Controllers Tests

	#Region Executing Exchange Server Tests
	write-host "Starting Exchange Servers..." -ForegroundColor Green
	if (Get-ExBoxStatus = $true)
	{
		$array_Exch_Checked = $null
		$array_EVS = $null
		if (($null -ne $clb_Step1_Nodes_List) -and ($intNodesTotal -gt 0))
		{
				write-host "Servers detected in Exchange Nodes box.  Merging Cluster Nodes into the Exchange Servers list."
				foreach ($item in $clb_Step1_Nodes_List.checkeditems)
				{
					$SplitItem = $item.split("~~")
					[array]$array_Exch_Checked += $SplitItem[2]
					[array]$array_EVS += $SplitItem[0]
				}
				$array_EVS = $array_EVS | select-object -Unique
				foreach ($item in $clb_Step1_Ex_List.checkeditems)
				{
					if ($array_EVS -notcontains $item){$array_Exch_Checked += $item.tostring()}
				}
		}
		else
		{
			write-host "No cluster nodes detected in Exchange Nodes box.  Using Exchange Servers check list box"
			foreach ($item in $clb_Step1_Ex_List.checkeditems)
			{
				[array]$array_Exch_Checked += $item.tostring()
			}
		}
		$array_Exch_Checked = $array_Exch_Checked | select-object -Unique

		if ($null -ne $array_Exch_Checked)
		{
			foreach ($server in $array_Exch_Checked)
			{
				$ping_reply = $null
				$ping = New-Object system.Net.NetworkInformation.Ping
				try
				{
					$ping_reply = $ping.send($server)
				}
				catch [system.Net.NetworkInformation.PingException]
				{
					write-host "Exception occured during Ping request to server: $server" -ForegroundColor Red
				}
				if (($ping_reply.status) -eq "Success") #This should check connections
#				if ((Test-Connection $server -count 2 -quiet) -eq $true) #This should check connections
				{
				    If ($chk_Ex_Win32_Bios.checked -eq $true)
						{Start-ExDCJob -server $server -job "Win32_BIOS" -JobType 0 -Location $location -JobScriptName "exch_w32_bios.ps1" -i $null -PSSession $null}
					If ($chk_Ex_Win32_ComputerSystem.checked -eq $true)
						{Start-ExDCJob -server $server -job "Win32_ComputerSystem" -JobType 0 -Location $location -JobScriptName "exch_w32_cs.ps1" -i $null -PSSession $null}
					If ($chk_EX_Win32_LogicalDisk.checked -eq $true)
						{Start-ExDCJob -server $server -job "Win32_LogicalDisk" -JobType 0 -Location $location -JobScriptName "exch_w32_ld.ps1" -i $null -PSSession $null}
					If ($chk_Ex_Win32_NetworkAdapter.checked -eq $true)
						{Start-ExDCJob -server $server -job "Win32_NetworkAdapter" -JobType 0 -Location $location -JobScriptName "exch_w32_na.ps1" -i $null -PSSession $null}
					If ($chk_Ex_Win32_NetworkAdapterConfig.checked -eq $true)
						{Start-ExDCJob -server $server -job "Win32_NetworkAdapterConfig" -JobType 0 -Location $location -JobScriptName "exch_w32_nac.ps1" -i $null -PSSession $null}
					If ($chk_Ex_Win32_OperatingSystem.checked -eq $true)
						{Start-ExDCJob -server $server -job "Win32_OperatingSystem" -JobType 0 -Location $location -JobScriptName "exch_w32_os.ps1" -i $null -PSSession $null}
				    If ($chk_Ex_Win32_PageFileUsage.checked -eq $true)
						{Start-ExDCJob -server $server -job "Win32_PageFileUsage" -JobType 0 -Location $location -JobScriptName "exch_w32_pfu.ps1" -i $null -PSSession $null}
				    If ($chk_Ex_Win32_PhysicalMemory.checked -eq $true)
						{Start-ExDCJob -server $server -job "Win32_PhysicalMemory" -JobType 0 -Location $location -JobScriptName "exch_w32_pm.ps1" -i $null -PSSession $null}
				    If ($chk_Ex_Win32_Processor.checked -eq $true)
						{Start-ExDCJob -server $server -job "Win32_Processor" -JobType 0 -Location $location -JobScriptName "exch_w32_proc.ps1" -i $null -PSSession $null}
				    If ($chk_Ex_Registry_Ex.checked -eq $true)
						{Start-ExDCJob -server $server -job "Registry - Exchange" -JobType 0 -Location $location -JobScriptName "exch_Reg_Ex.ps1" -i $null -PSSession $null}
					If ($chk_Ex_Registry_OS.checked -eq $true)
						{Start-ExDCJob -server $server -job "Registry - OS" -JobType 0 -Location $location -JobScriptName "exch_reg_os.ps1" -i $null -PSSession $null}
					If ($chk_Ex_Registry_Software.checked -eq $true)
						{Start-ExDCJob -server $server -job "Registry - Software" -JobType 0 -Location $location -JobScriptName "exch_reg_software.ps1" -i $null -PSSession $null}
				}
				else
				{
					$FailedPingOutput = ".\FailedPing_" + $append + ".txt"
					if ((Test-Path $FailedPingOutput) -eq $false)
			        {
				      new-item $FailedPingOutput -type file -Force
			        }
			        "* * * * * * * * * * * * * * * * * * * *"  | Out-File $FailedPingOutput -Force -Append
					write-host "---- Server $server failed Test-Connection for Exchange Server WMI functions" -ForegroundColor Red
					"Server: " + $server  | Out-File $FailedPingOutput -Force -Append
					"Failure: Ping reply to $server was not successful`n`n" | Out-File $FailedPingOutput -Force -Append
#					"Failure: Test-Connection $server -count 2 -quiet`n`n" | Out-File $FailedPingOutput -Force -Append
					$ErrorText = "ExDC " + "`n" + $server + "`n"
					$ErrorText += "Test-Connection failed"
					$ErrorLog = New-Object System.Diagnostics.EventLog('Application')
					$ErrorLog.MachineName = "."
					$ErrorLog.Source = "ExDC"
					try{$ErrorLog.WriteEntry($ErrorText,"Error", 400)} catch{}
				}
			}
		}
		else
		{
			write-host "No Exchange servers in the list are checked"
		}
	}
	else
	{
		write-host "---- No Exchange Server Functions selected"
	}
	#EndRegion Executing Exchange Server Tests

	#Region Executing Cluster Tests
	write-host "Starting Cluster..." -ForegroundColor Green
	If (Get-ClusterBoxStatus = $true)
	{
		if ($intNodesTotal -gt 0)
		{
			write-host "---- Cluster nodes detected in Exchange Nodes box.  Merging Cluster Nodes into the Exchange Servers list."
			foreach ($item in $clb_Step1_Nodes_List.checkeditems)
			{
				$SplitItem = $item.split("~~")
				[array]$array_Cluster_Checked += $SplitItem[2]
			}
			$array_Cluster_Checked = $array_Cluster_Checked | select-object -Unique

			if ($null -ne $array_Cluster_Checked)
			{
				foreach ($server in $array_Cluster_Checked)
				{
					$ping = New-Object system.Net.NetworkInformation.Ping
					try
					{
						$ping_reply = $ping.send($server)
					}
					catch [system.Net.NetworkInformation.PingException]
					{
						write-host "Exception occured during Ping request to server: $server" -ForegroundColor Red
					}
					if (($ping_reply.status) -eq "Success") #This should check connections
#					if ((Test-Connection $server -count 2 -quiet) -eq $true) #This should check connections
					{
					    If ($chk_Cluster_MSCluster_Node.checked -eq $true)
							{Start-ExDCJob -server $server -job "Cluster_MSCluster_Node" -JobType 0 -Location $location -JobScriptName "Cluster_MSCluster_Node.ps1" -i $null -PSSession $null}
					    If ($chk_Cluster_MSCluster_Network.checked -eq $true)
							{Start-ExDCJob -server $server -job "Cluster_MSCluster_Network" -JobType 0 -Location $location -JobScriptName "Cluster_MSCluster_Network.ps1" -i $null -PSSession $null}
					    If ($chk_Cluster_MSCluster_Resource.checked -eq $true)
							{Start-ExDCJob -server $server -job "Cluster_MSCluster_Resource" -JobType 0 -Location $location -JobScriptName "Cluster_MSCluster_Resource.ps1" -i $null -PSSession $null}
					    If ($chk_Cluster_MSCluster_ResourceGroup.checked -eq $true)
							{Start-ExDCJob -server $server -job "Cluster_MSCluster_ResourceGroup" -JobType 0 -Location $location -JobScriptName "Cluster_MSCluster_ResourceGroup.ps1" -i $null -PSSession $null}
					}
				}
			}
		}
		else
		{
			write-host "---- No cluster nodes detected in Exchange Nodes box."
		}
	}
	else
	{
		write-host "---- No Cluster Functions selected"
	}
	#EndRegion Executing Cluster Tests

	#Region Executing Exchange Organization Tests
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true) -or ($NoGUI -eq $true))
	{
		write-host "Starting Exchange Organization..." -ForegroundColor Green
		If (Get-ExOrgBoxStatus = $true)
		{
			# Save checked mailboxes to file for use by jobs
			$Mailbox_Checked_outputfile = ".\CheckedMailbox.txt"
			if ((Test-Path $Mailbox_Checked_outputfile) -eq $true)
			{
				Remove-Item $Mailbox_Checked_outputfile -Force
			}
			write-host "-- Building the checked mailbox list..."
			foreach ($item in $clb_Step1_Mailboxes_List.checkeditems)
			{
				$item.tostring() | out-file $Mailbox_Checked_outputfile -append -Force
			}

			If (Get-ExOrgMbxBoxStatus = $true)
			{
				# Avoid this path if we're not running mailbox tests
				# Splitting CheckedMailboxes file 10 times
				write-host "-- Splitting the list of checked mailboxes... "
				$File_Location = $location + "\CheckedMailbox.txt"
				If ((Test-Path $File_Location) -eq $false)
				{
					# Create empty Mailbox.txt file if not present
					write-host "No mailboxes appear to be selected.  Mailbox tests will produce no output." -ForegroundColor Red
					"" | Out-File $File_Location
				}
				$CheckedMailbox = [System.IO.File]::ReadAllLines($File_Location)
				$CheckedMailboxCount = $CheckedMailbox.count
				$CheckedMailboxCountSplit = [int]$CheckedMailboxCount/10
				if ((Test-Path ".\CheckedMailbox.Set1.txt") -eq $true) {Remove-Item ".\CheckedMailbox.Set1.txt" -Force}
				if ((Test-Path ".\CheckedMailbox.Set2.txt") -eq $true) {Remove-Item ".\CheckedMailbox.Set2.txt" -Force}
				if ((Test-Path ".\CheckedMailbox.Set3.txt") -eq $true) {Remove-Item ".\CheckedMailbox.Set3.txt" -Force}
				if ((Test-Path ".\CheckedMailbox.Set4.txt") -eq $true) {Remove-Item ".\CheckedMailbox.Set4.txt" -Force}
				if ((Test-Path ".\CheckedMailbox.Set5.txt") -eq $true) {Remove-Item ".\CheckedMailbox.Set5.txt" -Force}
				if ((Test-Path ".\CheckedMailbox.Set6.txt") -eq $true) {Remove-Item ".\CheckedMailbox.Set6.txt" -Force}
				if ((Test-Path ".\CheckedMailbox.Set7.txt") -eq $true) {Remove-Item ".\CheckedMailbox.Set7.txt" -Force}
				if ((Test-Path ".\CheckedMailbox.Set8.txt") -eq $true) {Remove-Item ".\CheckedMailbox.Set8.txt" -Force}
				if ((Test-Path ".\CheckedMailbox.Set9.txt") -eq $true) {Remove-Item ".\CheckedMailbox.Set9.txt" -Force}
				if ((Test-Path ".\CheckedMailbox.Set10.txt") -eq $true) {Remove-Item ".\CheckedMailbox.Set10.txt" -Force}
				For ($Count = 0;$Count -lt ($CheckedMailboxCountSplit);$Count++)
					{$CheckedMailbox[$Count] | Out-File ".\CheckedMailbox.Set1.txt" -Append -Force}
				For (;$Count -lt (2*$CheckedMailboxCountSplit);$Count++)
					{$CheckedMailbox[$Count] | Out-File ".\CheckedMailbox.Set2.txt" -Append -Force}
				For (;$Count -lt (3*$CheckedMailboxCountSplit);$Count++)
					{$CheckedMailbox[$Count] | Out-File ".\CheckedMailbox.Set3.txt" -Append -Force}
				For (;$Count -lt (4*$CheckedMailboxCountSplit);$Count++)
					{$CheckedMailbox[$Count] | Out-File ".\CheckedMailbox.Set4.txt" -Append -Force}
				For (;$Count -lt (5*$CheckedMailboxCountSplit);$Count++)
					{$CheckedMailbox[$Count] | Out-File ".\CheckedMailbox.Set5.txt" -Append -Force}
				For (;$Count -lt (6*$CheckedMailboxCountSplit);$Count++)
					{$CheckedMailbox[$Count] | Out-File ".\CheckedMailbox.Set6.txt" -Append -Force}
				For (;$Count -lt (7*$CheckedMailboxCountSplit);$Count++)
					{$CheckedMailbox[$Count] | Out-File ".\CheckedMailbox.Set7.txt" -Append -Force}
				For (;$Count -lt (8*$CheckedMailboxCountSplit);$Count++)
					{$CheckedMailbox[$Count] | Out-File ".\CheckedMailbox.Set8.txt" -Append -Force}
				For (;$Count -lt (9*$CheckedMailboxCountSplit);$Count++)
					{$CheckedMailbox[$Count] | Out-File ".\CheckedMailbox.Set9.txt" -Append -Force}
				For (;$Count -lt (10*$CheckedMailboxCountSplit);$Count++)
					{$CheckedMailbox[$Count] | Out-File ".\CheckedMailbox.Set10.txt" -Append -Force}
			}

			# First we start the jobs that query the organization instead of the Exchange server
			#Region ExOrg Non-server Functions
			If ($chk_Org_Get_AcceptedDomain.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-AcceptedDomain" -JobType 1 -Location $location -JobScriptName "ExOrg_GetAcceptedDomain.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_ActiveSyncPolicy.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-ActiveSyncMailboxPolicy" -JobType 1 -Location $location -JobScriptName "ExOrg_GetActiveSyncMbxPolicy.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_AddressBookPolicy.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-AddressBookPolicy" -JobType 1 -Location $location -JobScriptName "ExOrg_GetAddressBookPolicy.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_AddressList.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-AddressList" -JobType 1 -Location $location -JobScriptName "ExOrg_GetAddressList.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_AdPermission.checked -eq $true)
			{
				For ($i = 1;$i -lt 11;$i++)
				{Start-ExDCJob -server $server -job "Get-AdPermission - Set $i" -JobType 1 -Location $location -JobScriptName "ExOrg_GetAdPermission.ps1" -i $i -PSSession $session_0}
			}
			If ($chk_Org_Get_AdSite.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-AdSite" -JobType 1 -Location $location -JobScriptName "ExOrg_GetAdSite.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_AdSiteLink.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-AdSiteLink" -JobType 1 -Location $location -JobScriptName "ExOrg_GetAdSiteLink.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_CASMailbox.checked -eq $true)
			{
				For ($i = 1;$i -lt 11;$i++)
				{Start-ExDCJob -server $server -job "Get-CASMailbox - Set $i" -JobType 1 -Location $location -JobScriptName "ExOrg_GetCASMailbox.ps1" -i $i -PSSession $session_0}
			}
			If ($chk_Org_Get_ClientAccessServer.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-ClientAccessServer" -JobType 1 -Location $location -JobScriptName "ExOrg_GetClientAccessSvr.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_ContentFilterConfig.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-ContentFilterConfig" -JobType 1 -Location $location -JobScriptName "ExOrg_GetContentFilterConfig.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_DistributionGroup.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-DistributionGroup" -JobType 1 -Location $location -JobScriptName "ExOrg_GetDistributionGroup.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_DynamicDistributionGroup.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-DynamicDistributionGroup" -JobType 1 -Location $location -JobScriptName "ExOrg_GetDynamicDistributionGroup.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_EmailAddressPolicy.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-EmailAddressPolicy" -JobType 1 -Location $location -JobScriptName "ExOrg_GetEmailAddressPolicy.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_ExchangeServer.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-ExchangeServer" -JobType 1 -Location $location -JobScriptName "ExOrg_GetExchSvr.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_Mailbox.checked -eq $true)
			{
				For ($i = 1;$i -lt 11;$i++)
				{Start-ExDCJob -server $server -job "Get-Mailbox - Set $i" -JobType 1 -Location $location -JobScriptName "ExOrg_GetMbx.ps1" -i $i -PSSession $session_0}
			}
            If ($chk_Org_Get_User.checked -eq $true)
			{
				For ($i = 1;$i -lt 11;$i++)
				{Start-ExDCJob -server $server -job "Get-User - Set $i" -JobType 1 -Location $location -JobScriptName "ExOrg_GetUser.ps1" -i $i -PSSession $session_0}
			}
			If ($chk_Org_Get_MailboxFolderStatistics.checked -eq $true)
			{
				For ($i = 1;$i -lt 11;$i++)
				{Start-ExDCJob -server $server -job "Get-MailboxFolderStatistics - Set $i" -JobType 1 -Location $location -JobScriptName "ExOrg_GetMbxFolderStatistics.ps1" -i $i -PSSession $session_0}
			}
			If ($chk_Org_Get_MailboxPermission.checked -eq $true)
			{
				For ($i = 1;$i -lt 11;$i++)
				{Start-ExDCJob -server $server -job "Get-MailboxPermission - Set $i" -JobType 1 -Location $location -JobScriptName "ExOrg_GetMbxPermission.ps1" -i $i -PSSession $session_0}
			}
			If ($chk_Org_Get_MailboxServer.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-MailboxServer" -JobType 1 -Location $location -JobScriptName "ExOrg_GetMbxSvr.ps1" -i $null -PSSession $session_0}
		    If ($chk_Org_Get_MailboxStatistics.checked -eq $true)
			{
				For ($i = 1;$i -lt 11;$i++)
				{Start-ExDCJob -server $server -job "Get-MailboxStatistics - Set $i" -JobType 1 -Location $location -JobScriptName "ExOrg_GetMbxStatistics.ps1" -i $i -PSSession $session_0}
			}
			If ($chk_Org_Get_OfflineAddressBook.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-OfflineAddressBook" -JobType 1 -Location $location -JobScriptName "ExOrg_GetOfflineAddressBook.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_OrgConfig.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-OrganizationConfig" -JobType 1 -Location $location -JobScriptName "ExOrg_GetOrgConfig.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_OutlookAnywhere.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-OutlookAnywhere" -JobType 1 -Location $location -JobScriptName "ExOrg_GetOutlookAnywhere.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_OwaMailboxPolicy.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-OwaMailboxPolicy" -JobType 1 -Location $location -JobScriptName "ExOrg_GetOwaMailboxPolicy.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_ReceiveConnector.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-ReceiveConnector" -JobType 1 -Location $location -JobScriptName "ExOrg_GetReceiveConnector.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_RemoteDomain.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-RemoteDomain" -JobType 1 -Location $location -JobScriptName "ExOrg_GetRemoteDomain.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_Rbac.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-Rbac" -JobType 1 -Location $location -JobScriptName "ExOrg_GetRbac.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_RetentionPolicy.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-RetentionPolicy" -JobType 1 -Location $location -JobScriptName "ExOrg_GetRetentionPolicy.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_RetentionPolicyTag.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-RetentionPolicyTag" -JobType 1 -Location $location -JobScriptName "ExOrg_GetRetentionPolicyTag.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_RoutingGroupConnector.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-RoutingGroupConnector" -JobType 1 -Location $location -JobScriptName "ExOrg_GetRoutingGroupConnector.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_SendConnector.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-SendConnector" -JobType 1 -Location $location -JobScriptName "ExOrg_GetSendConnector.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_TransportConfig.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-TransportConfig" -JobType 1 -Location $location -JobScriptName "ExOrg_GetTransportConfig.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_TransportRule.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-TransportRule" -JobType 1 -Location $location -JobScriptName "ExOrg_GetTransportRule.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_TransportServer.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-TransportServer" -JobType 1 -Location $location -JobScriptName "ExOrg_GetTransportSvr.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_UmAutoAttendant.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-UmAutoAttendant" -JobType 1 -Location $location -JobScriptName "ExOrg_GetUmAutoAttendant.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_UmDialPlan.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-UmDialPlan" -JobType 1 -Location $location -JobScriptName "ExOrg_GetUmDialPlan.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_UmIpGateway.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-UmIpGateway" -JobType 1 -Location $location -JobScriptName "ExOrg_GetUmIpGateway.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_UmMailbox.checked -eq $true)
			{
				For ($i = 1;$i -lt 11;$i++)
				{Start-ExDCJob -server $server -job "Get-UmMailbox - Set $i" -JobType 1 -Location $location -JobScriptName "ExOrg_GetUmMailbox.ps1" -i $i -PSSession $session_0}
			}
#			If ($chk_Org_Get_UmMailboxConfiguration.checked -eq $true)
#			{
#				For ($i = 1;$i -lt 11;$i++)
#				{Start-ExDCJob -server $server -job "Get-UmMailboxConfiguration - Set $i" -JobType 1 -Location $location -JobScriptName "ExOrg_GetUmMailboxConfiguration.ps1" -i $i -PSSession $session_0}
#			}
#			If ($chk_Org_Get_UmMailboxPin.checked -eq $true)
#			{
#				For ($i = 1;$i -lt 11;$i++)
#				{Start-ExDCJob -server $server -job "Get-UmMailboxPin - Set $i" -JobType 1 -Location $location -JobScriptName "ExOrg_GetUmMailboxPin.ps1" -i $i -PSSession $session_0}
#			}
			If ($chk_Org_Get_UmMailboxPolicy.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-UmMailboxPolicy" -JobType 1 -Location $location -JobScriptName "ExOrg_GetUmMailboxPolicy.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_UmServer.checked -eq $true)
				{Start-ExDCJob -server $server -job "Get-UmServer" -JobType 1 -Location $location -JobScriptName "ExOrg_GetUmSvr.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Quota.checked -eq $true)
			{
				For ($i = 1;$i -lt 11;$i++)
				{Start-ExDCJob -server $server -job "Quota - Set $i" -JobType 1 -Location $location -JobScriptName "ExOrg_Quota.ps1" -i $i -PSSession $session_0}
			}
			If ($chk_Org_Get_AdminGroups.checked -eq $true)
				{Start-ExDCJob -server $server -job "Misc - Get Admin Groups" -JobType 1 -Location $location -JobScriptName "ExOrg_Misc_AdminGroups.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_Fsmo.checked -eq $true)
				{Start-ExDCJob -server $server -job "Misc - Get FSMO Roles" -JobType 1 -Location $location -JobScriptName "ExOrg_Misc_Fsmo.ps1" -i $null -PSSession $session_0}
			If ($chk_Org_Get_ExchangeServerBuilds.checked -eq $true)
				{Start-ExDCJob -server $server -job "Misc - Get Exchange Server Builds" -JobType 1 -Location $location -JobScriptName "ExOrg_Misc_ExchangeBuilds.ps1" -i $null -PSSession $session_0}
			If (($Exchange2010 -eq $true) -or ($Exchange2013 -eq $true) -or ($Exchange2016 -eq $true)) 
			{
				# Exchange 2010+ only cmdlet
				If ($chk_Org_Get_ActiveSyncDevice.checked -eq $true)
					{Start-ExDCJob -server $server -job "Get-ActiveSyncDevice" -JobType 1 -Location $location -JobScriptName "ExOrg_GetActiveSyncDevice.ps1" -i $null -PSSession $session_0}
				# Exchange 2010+ only cmdlet
				If ($chk_Org_Get_CalendarProcessing.checked -eq $true)
				{
					For ($i = 1;$i -lt 11;$i++)
					{Start-ExDCJob -server $server -job "Get-CalendarProcessing - Set $i" -JobType 1 -Location $location -JobScriptName "ExOrg_GetCalendarProcessing.ps1" -i $i -PSSession $session_0}
				}
				# Exchange 2010+ only cmdlet
				If ($chk_Org_Get_AvailabilityAddressSpace.checked -eq $true)
					{Start-ExDCJob -server $server -job "Get-AvailabilityAddressSpace" -JobType 1 -Location $location -JobScriptName "ExOrg_GetAvailabilityAddressSpace.ps1" -i $null -PSSession $session_0}
				# Exchange 2010+ only cmdlet
				If ($chk_Org_Get_ClientAccessArray.checked -eq $true)
					{Start-ExDCJob -server $server -job "Get-ClientAccessArray" -JobType 1 -Location $location -JobScriptName "ExOrg_GetClientAccessArray.ps1" -i $null -PSSession $session_0}
				# Exchange 2010+ only cmdlet
				If ($chk_Org_Get_DatabaseAvailabilityGroup.checked -eq $true)
					{Start-ExDCJob -server $server -job "Get-DatabaseAvailabilityGroup" -JobType 1 -Location $location -JobScriptName "ExOrg_GetDatabaseAvailabilityGroup.ps1" -i $null -PSSession $session_0}
				# Exchange 2010+ only cmdlet
				If ($chk_Org_Get_DAGNetwork.checked -eq $true)
					{Start-ExDCJob -server $server -job "Get-DAGNetwork" -JobType 1 -Location $location -JobScriptName "ExOrg_GetDatabaseAvailabilityGroupNetwork.ps1" -i $null -PSSession $session_0}
				# Exchange 2010+ only cmdlet
				If ($chk_Org_Get_RPCClientAccess.checked -eq $true)
					{Start-ExDCJob -server $server -job "Get-RPCClientAccess" -JobType 1 -Location $location -JobScriptName "ExOrg_GetRpcClientAccess.ps1" -i $null -PSSession $session_0}
				# Exchange 2010+ only cmdlet
				If ($chk_Org_Get_ThrottlingPolicy.checked -eq $true)
					{Start-ExDCJob -server $server -job "Get-ThrottlingPolicy" -JobType 1 -Location $location -JobScriptName "ExOrg_GetThrottlingPolicy.ps1" -i $null -PSSession $session_0}
			}
			#EndRegion ExOrg Non-Server Functions

			# Then start the ExOrg jobs that will use the $server variable
			#Region ExOrg Server Functions
			if (Get-ExOrgServerStatus -eq $true)
			{
				$array_ExOrg_Server_Checked = $null
				foreach ($item in $clb_Step1_Ex_List.checkeditems)
				{
					[array]$array_ExOrg_Server_Checked += $item.tostring()
				}
				if ($null -ne $array_ExOrg_Server_Checked)
				{
					foreach ($server in $array_ExOrg_Server_Checked)
					{
						$ping = New-Object system.Net.NetworkInformation.Ping
						try
						{
							$ping_reply = $ping.send($server)
						}
						catch [system.Net.NetworkInformation.PingException]
						{
							write-host "Exception occured during Ping request to server: $server" -ForegroundColor Red
						}
						if (($ping_reply.status) -eq "Success") #This should check connections
	#					if ((Test-Connection $server -count 2 -quiet) -eq $true) #This line caused Powershell to silently exit against E2k7 servers
						{
							$ServerInfo = Get-ExchangeServer -identity $server
							if ($ServerInfo.ServerRole -match "Mailbox")
							{
								$PFServerInfo = Get-PublicFolderDatabase -server $server
							}
							else
							{
								$PFServerInfo = $null
							}

							# Get-MailboxDatabase and Get-PublicFolderDatabase can be run against E2k3
						    If (($chk_Org_Get_MailboxDatabase.checked -eq $true) -and ($null -ne ($ServerInfo.ServerRole -match "Mailbox")))
								{Start-ExDCJob -server $server -job "Get-MailboxDatabase" -JobType 1 -Location $location -JobScriptName "ExOrg_GetMbxDb.ps1" -i $null -PSSession $session_0}
							If (($chk_Org_Get_PublicFolderDatabase.checked -eq $true) -and ($null -ne $PFServerInfo))
								{Start-ExDCJob -server $server -job "Get-PublicFolderDatabase" -JobType 1 -Location $location -JobScriptName "ExOrg_GetPublicFolderDb.ps1" -i $null -PSSession $session_0}
							# These cmdlets are E2k7 only
							if ((($ServerInfo.IsExchange2007OrLater -eq $true) -and ($ServerInfo.IsE14OrLater -eq $null)) -or
								(($ServerInfo.IsExchange2007OrLater -eq $true) -and ($ServerInfo.IsE14OrLater -eq $false)))
								{Start-ExDCJob -server $server -job "Get-StorageGroup" -JobType 1 -Location $location -JobScriptName "ExOrg_GetStorageGroup.ps1" -i $null -PSSession $session_0}

							# These cmdlets will fail against pre-E2k7 servers
							if ($ServerInfo.IsExchange2007OrLater -eq $true)
							{
								If ($chk_Org_Get_ActiveSyncVirtualDirectory.checked -eq $true)
									{Start-ExDCJob -server $server -job "Get-ActiveSyncVirtualDirectoryjob" -JobType 1 -Location $location -JobScriptName "ExOrg_GetActiveSyncVD.ps1" -i $null -PSSession $session_0}
								If ($chk_Org_Get_AutodiscoverVirtualDirectory.checked -eq $true)
									{Start-ExDCJob -server $server -job "Get-AutodiscoverVirtualDirectory" -JobType 1 -Location $location -JobScriptName "ExOrg_GetAutodiscoverVirtualDirectory.ps1" -i $null -PSSession $session_0}
								If ($chk_Org_Get_OABVirtualDirectory.checked -eq $true)
									{Start-ExDCJob -server $server -job "Get-OABVirtualDirectory" -JobType 1 -Location $location -JobScriptName "ExOrg_GetOABVD.ps1" -i $null -PSSession $session_0}
								If ($chk_Org_Get_OWAVirtualDirectory.checked -eq $true)
									{Start-ExDCJob -server $server -job "Get-OWAVirtualDirectory" -JobType 1 -Location $location -JobScriptName "ExOrg_GetOWAVD.ps1" -i $null -PSSession $session_0}
								If (($chk_Org_Get_PublicFolder.checked -eq $true) -and ($null -ne $PFServerInfo))
									{Start-ExDCJob -server $server -job "Get-PublicFolder" -JobType 1 -Location $location -JobScriptName "ExOrg_GetPublicFolder.ps1" -i $null -PSSession $session_0}
							    If (($chk_Org_Get_PublicFolderStatistics.checked -eq $true) -and ($null -ne $PFServerInfo))
									{Start-ExDCJob -server $server -job "Get-PublicFolderStatistics" -JobType 1 -Location $location -JobScriptName "ExOrg_GetPublicFolderStats.ps1" -i $null -PSSession $session_0}
								If ($chk_Org_Get_WebServicesVirtualDirectory.checked -eq $true)
									{Start-ExDCJob -server $server -job "Get-WebServicesVirtualDirectory" -JobType 1 -Location $location -JobScriptName "ExOrg_GetWebServVD.ps1" -i $null -PSSession $session_0}
							}

							# These cmdlets will fail against pre-E14 servers
							if ($ServerInfo.IsE14OrLater -eq $true)
							{
								# Exchange 2010+ only cmdlet - E2k7 does not support -server parameter
								If ($chk_Org_Get_ExchangeCertificate.checked -eq $true)
									{Start-ExDCJob -server $server -job "Get-ExchangeCertificate" -JobType 1 -Location $location -JobScriptName "ExOrg_GetExchangeCertificate.ps1" -i $null -PSSession $session_0}
								# Exchange 2010+ only cmdlet
								If ($chk_Org_Get_ECPVirtualDirectory.checked -eq $true)
									{Start-ExDCJob -server $server -job "Get-ECPVirtualDirectory" -JobType 1 -Location $location -JobScriptName "ExOrg_GetEcpVirtualDirectory.ps1" -i $null -PSSession $session_0}
								# Exchange 2010+ only cmdlet
								If (($chk_Org_Get_MailboxDatabaseCopyStatus.checked -eq $true) -and (($ServerInfo.ServerRole -match "Mailbox")))
									{Start-ExDCJob -server $server -job "Get-MailboxDatabaseCopyStatus" -JobType 1 -Location $location -JobScriptName "ExOrg_GetMailboxDatabaseCopyStatus.ps1" -i $null -PSSession $session_0}
								# Exchange 2010+ only cmdlet
								If ($chk_Org_Get_PowershellVirtualDirectory.checked -eq $true)
									{Start-ExDCJob -server $server -job "Get-PowershellVirtualDirectory" -JobType 1 -Location $location -JobScriptName "ExOrg_GetPowershellVD.ps1" -i $null -PSSession $session_0}
							}
						}
						else
						{
							$FailedPingOutput = ".\FailedPing_" + $append + ".txt"
							if ((Test-Path $FailedPingOutput) -eq $false)
					        {
						      new-item $FailedPingOutput -type file -Force
					        }
					        "* * * * * * * * * * * * * * * * * * * *"  | Out-File $FailedPingOutput -Force -Append
							write-host "---- Server $server failed Test-Connection  for Exchange Organization functions" -ForegroundColor Red
							"Server: " + $server  | Out-File $FailedPingOutput -Force -Append
							"Failure: Ping reply to $server was not successful`n`n" | Out-File $FailedPingOutput -Force -Append
	#						"Failure: Test-Connection $server -count 2 -quiet`n`n" | Out-File $FailedPingOutput -Force -Append
							$ErrorText = "ExDC " + "`n" + $server + "`n"
							$ErrorText += "Test-Connection failed"
							$ErrorLog = New-Object System.Diagnostics.EventLog('Application')
							$ErrorLog.MachineName = "."
							$ErrorLog.Source = "ExDC"
							Try{$ErrorLog.WriteEntry($ErrorText,"Error", 400)} catch{}
						}
					}
				}
				else
				{
					write-host "No Exchange servers in the list are checked"
				}
			}
			#EndRegion ExOrg Server Functions
		}
		else
		{
			write-host "---- No Exchange Organization Functions selected"
		}
	}
	#EndRegion Executing Exchange Organization Tests


	# Delay changing status to Idle until all jobs have finished
	Update-ExDCJobCount 1 15
	Remove-Item	".\RunningJobs.txt"
	# Remove Failed Jobs
	$colJobsFailed = @(Get-Job -State Failed)
	foreach ($objJobsFailed in $colJobsFailed)
	{
		if ($objJobsFailed.module -like "__DynamicModule*")
		{
			Remove-Job -Id $objJobsFailed.id
		}
		else
		{
            write-host "---- Failed job " $objJobsFailed.name -ForegroundColor Red
			$FailedJobOutput = ".\FailedJobs_" + $append + ".txt"
            if ((Test-Path $FailedJobOutput) -eq $false)
	        {
		      new-item $FailedJobOutput -type file -Force
	        }
	        "* * * * * * * * * * * * * * * * * * * *"  | Out-File $FailedJobOutput -Force -Append
            "Job Name: " + $objJobsFailed.name | Out-File $FailedJobOutput -Force -Append
	        "Job State: " + $objJobsFailed.state | Out-File $FailedJobOutput -Force	-Append
            if ($null -ne ($objJobsFailed.childjobs[0]))
            {
	           $objJobsFailed.childjobs[0].output | format-list | Out-File $FailedJobOutput -Force -Append
	           $objJobsFailed.childjobs[0].warning | format-list | Out-File $FailedJobOutput -Force -Append
	           $objJobsFailed.childjobs[0].error | format-list | Out-File $FailedJobOutput -Force -Append
			}
            $ErrorText = $objJobsFailed.name + "`n"
			$ErrorText += "Job failed"
			$ErrorLog = New-Object System.Diagnostics.EventLog('Application')
			$ErrorLog.MachineName = "."
			$ErrorLog.Source = "ExDC"
			Try{$ErrorLog.WriteEntry($ErrorText,"Error", 500)} catch{}
			Remove-Job -Id $objJobsFailed.id
		}
	}
	write-host "Restoring ExDC Form to normal." -ForegroundColor Green
	$form1.WindowState = "normal"
	$btn_Step3_Execute.enabled = $true
	$status_Step3.Text = "Step 3 Status: Idle"
	write-host "Step 3 jobs finished"
    Get-Job | Remove-Job -Force
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "ExDC"
	try{$EventLog.WriteEntry("Ending ExDC Step 3","Information", 31)} catch{}
    Stop-Transcript
}

Function Start-NoGUI()
{
	#Write-Host "Entered Start-NoGUI function"
	# If -NoGUI is set, then we're going to handle everything automatically
	If ($NoGUI -eq $true)
	{
		# Populate the targets
		Import-TargetsDc | Out-Null
		Import-TargetsEx | Out-Null
		Import-TargetsNodes | Out-Null
		If ($INI_ExOrg -ne "")
		{
			$File_Location = $location + "\mailbox.txt"
		    if ((Test-Path $File_Location) -eq $true)
			{
				$EventLog = New-Object System.Diagnostics.EventLog('Application')
				$EventLog.MachineName = "."
				$EventLog.Source = "ExDC"
				try{$EventLog.WriteEntry("Starting ExDC Step 1 - Populate","Information", 10)} catch{}
			    $array_Mailboxes = @(([System.IO.File]::ReadAllLines($File_Location)) | sort-object -Unique)
				$global:intMailboxTotal = 0
			    $clb_Step1_Mailboxes_List.items.clear()
				foreach ($member_Mailbox in $array_Mailboxes | where-object {$_ -ne ""})
			    {
			        $clb_Step1_Mailboxes_List.items.add($member_Mailbox) | Out-Null
					$global:intMailboxTotal++
			    }
				For ($i=0;$i -le ($intMailboxTotal - 1);$i++)
				{
					$clb_Step1_Mailboxes_List.SetItemChecked($i,$true)
				}
				$EventLog = New-Object System.Diagnostics.EventLog('Application')
				$EventLog.MachineName = "."
				$EventLog.Source = "ExDC"
				try{$EventLog.WriteEntry("Ending ExDC Step 1 - Populate","Information", 11)} catch{}
				#$txt_MailboxesTotal.Text = "Mailbox count = " + $intMailboxTotal
				#$txt_MailboxesTotal.visible = $true
			    #$status_Step1.Text = "Step 2 Status: Idle"
			}
			else
			{
				write-host	"The file mailbox.txt is not present.  Run Discover to create the file."
				#$status_Step1.Text = "Step 1 Status: Failed - mailbox.txt file not found.  Run Discover to create the file."
			}
		}
		# Click the Execute Button
		Start-Execute
		
		# Rename the folder
		If ($NoGUIOutputFolder -ne "")
		{
			$Newfolder = $NoGUIOutputFolder
			Rename-Item $location"\output" $location"\"$NoGUIOutputFolderOutputFolder
		}
		else
		{
			$ticks = (Get-Date).ticks
			$NewFolder = "output-$ticks"
			Rename-Item $location"\output" $location"\"$NewFolder
		}
	}
	exit
}



Function Disable-AllTargetsButtons()
{
    $btn_Step1_DC_Discover.enabled = $false
    $btn_Step1_DC_Populate.enabled = $false
    $btn_Step1_Ex_Discover.enabled = $false
    $btn_Step1_Ex_Populate.enabled = $false
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2003orEarlier -eq $true) -or ($NoEMS -eq $true))
		{$btn_Step1_Nodes_Discover.enabled = $false}
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2003orEarlier -eq $true) -or ($NoEMS -eq $true))
		{$btn_Step1_Nodes_Populate.enabled = $false}
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true))
		{$btn_Step1_Mailboxes_Discover.enabled = $false}
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true))
		{$btn_Step1_Mailboxes_Populate.enabled = $false}
}

Function Enable-AllTargetsButtons()
{
    $btn_Step1_DC_Discover.enabled = $true
    $btn_Step1_DC_Populate.enabled = $true
    $btn_Step1_Ex_Discover.enabled = $true
    $btn_Step1_Ex_Populate.enabled = $true
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2003orEarlier -eq $true) -or ($NoEMS -eq $true))
		{$btn_Step1_Nodes_Discover.enabled = $true}
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2003orEarlier -eq $true) -or ($NoEMS -eq $true))
		{$btn_Step1_Nodes_Populate.enabled = $true}
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true))
		{$btn_Step1_Mailboxes_Discover.enabled = $true}
	if (($Exchange2007Powershell -eq $true) -or ($Exchange2010Powershell -eq $true))
		{$btn_Step1_Mailboxes_Populate.enabled = $true}
}

Function Limit-ExDCJob
{
	Param(	[int]$JobThrottleMaxJobs,`
			[int]$JobThrottlePolling,`
			[string]$PsSession
			)

	# This function is only called during Start-ExDCJob to prevent jobs from starting prematurely
	# Remove Failed Jobs
	$colJobsFailed = @(Get-Job -State Failed)
	foreach ($objJobsFailed in $colJobsFailed)
	{
		if ($objJobsFailed.module -like "__DynamicModule*")
		{
			Remove-Job -Id $objJobsFailed.id
		}
		else
		{
            write-host "---- Failed job " $objJobsFailed.name -ForegroundColor Red
			$FailedJobOutput = ".\FailedJobs_" + $append + ".txt"
            if ((Test-Path $FailedJobOutput) -eq $false)
	        {
		      new-item $FailedJobOutput -type file -Force
	        }
	        "* * * * * * * * * * * * * * * * * * * *"  | Out-File $FailedJobOutput -Force -Append
            "Job Name: " + $objJobsFailed.name | Out-File $FailedJobOutput -Force -Append
	        "Job State: " + $objJobsFailed.state | Out-File $FailedJobOutput -Force	-Append
            if ($null -ne ($objJobsFailed.childjobs[0]))
            {
	           $objJobsFailed.childjobs[0].output | format-list | Out-File $FailedJobOutput -Force -Append
	           $objJobsFailed.childjobs[0].warning | format-list | Out-File $FailedJobOutput -Force -Append
	           $objJobsFailed.childjobs[0].error | format-list | Out-File $FailedJobOutput -Force -Append
			}
            $ErrorText = $objJobsFailed.name + "`n"
			$ErrorText += "Job failed"
			$ErrorLog = New-Object System.Diagnostics.EventLog('Application')
			$ErrorLog.MachineName = "."
			$ErrorLog.Source = "ExDC"
			Try{$ErrorLog.WriteEntry($ErrorText,"Error", 500)} catch{}
			Remove-Job -Id $objJobsFailed.id
		}
	}

	# Remove Completed Jobs
	$colJobsCompleted = @((Get-Job -State completed) | where-object {$null -ne $_.childjobs})
	foreach ($objJobsCompleted in $colJobsCompleted)
	{
		Remove-Job -Id $objJobsCompleted.id
		write-host "---- Finished job " $objJobsCompleted.name -ForegroundColor Green
	}

	# Get Running Jobs
    $colJobsRunning = @((Get-Job -State Running) | where-object {$_.Module -ne "__DynamicModule*"})
	if ((Test-Path ".\RunningJobs.txt") -eq $false)
	{
		new-item ".\RunningJobs.txt" -type file -Force
	}
	$RunningJobsOutput = ""
	$Now = Get-Date
	foreach ($objJobsRunning in $colJobsRunning)
	{
		$JobPID = $objJobsRunning.childjobs[0].output[0]
		if ($null -ne $JobPID)
		{
			# Pass the variable assignment as a condition to reduce timing issues
			if(($JobStartTime = ((Get-Process | where-object {$_.id -eq $JobPID}).starttime)) -ne $null)
			{
				$JobRunningTime = [int](($Now - $JobStartTime).TotalMinutes)
				if ((($objJobsRunning.childjobs[0].output[1] -eq "WMI") -and ($JobRunningTime -gt ($intWMIJobTimeout/60))) `
					-or (($objJobsRunning.childjobs[0].output[1] -eq "ExOrg") -and ($JobRunningTime -gt ($intExchJobTimeout/60))))
				{
					try
					{
						(Get-Process | where-object {$_.id -eq $JobPID}).kill()
						write-host "Timer expired.  Killing job process $JobPID - " + $objJobsRunning.name -ForegroundColor Red
						$ErrorText = $objJobsRunning.name + "`n"
						$ErrorText += "Process $JobPID killed`n"
						if ($objJobsRunning.childjobs[0].output[1] -eq "WMI") {$ErrorText += "Timeout $intWMIJobTimeout seconds exceeded"}
						if ($objJobsRunning.childjobs[0].output[1] -eq "ExOrg") {$ErrorText += "Timeout $intExchJobTimeout seconds exceeded"}
						$ErrorLog = New-Object System.Diagnostics.EventLog('Application')
						$ErrorLog.MachineName = "."
						$ErrorLog.Source = "ExDC"
						Try{$ErrorLog.WriteEntry($ErrorText,"Error", 600)} catch{}
					}
					catch [System.Management.Automation.MethodInvocationException]
					{
						write-host "`tMethodInvocationException occured during Kill request for process $JobPID" -ForegroundColor Red
					}
					catch [System.Management.Automation.RuntimeException]
					{
						write-host "`tRuntimeException occured during Kill request for process $JobPID" -ForegroundColor Red
					}
				}
				$RunningJobsOutput += "Job Name: " + $objJobsRunning.name + "`n"
				$RunningJobsOutput += "Job State: " + $objJobsRunning.State + "`n"
				$RunningJobsOutput += "Job process PID: " + $JobPID + "`n"
				$RunningJobsOutput += "Job process time running: " +  $JobRunningTime + " min"
				$RunningJobsOutput += "`n`n"
			}
		}
	}
	$RunningJobsOutput | Out-File ".\RunningJobs.txt" -Force

	# Limit Jobs based on PSSessions instead of job counts
	$intRunningJobs = Get-PsSessionCount -CurrentUser $CurrentUser -PsSession $PsSession -ResourceUri $ResourceUri
	#$intRunningJobs = $colJobsRunning.count
	if ($intRunningJobs -eq $null)
	{
		$intRunningJobs = "0"
	}

	do
	{
        ## Repeat bulk of function code to prevent recursive loop
        ##      and the dreaded System.Management.Automation.ScriptCallDepthException:
        ##      The script failed due to call depth overflow.
        ##      The call depth reached 1001 and the maximum is 1000.

            # Remove Failed Jobs
	        $colJobsFailed = @(Get-Job -State Failed)
	        foreach ($objJobsFailed in $colJobsFailed)
            {
		        if ($objJobsFailed.module -like "__DynamicModule*")
		        {
			        Remove-Job -Id $objJobsFailed.id
		        }
		        else
		        {
                    write-host "---- Failed job " $objJobsFailed.name -ForegroundColor Red
			        $FailedJobOutput = ".\FailedJobs_" + $append + ".txt"
                    if ((Test-Path $FailedJobOutput) -eq $false)
	                {
		              new-item $FailedJobOutput -type file -Force
	                }
	                "* * * * * * * * * * * * * * * * * * * *"  | Out-File $FailedJobOutput -Force -Append
                    "Job Name: " + $objJobsFailed.name | Out-File $FailedJobOutput -Force -Append
	                "Job State: " + $objJobsFailed.state | Out-File $FailedJobOutput -Force	-Append
                    if ($null -ne ($objJobsFailed.childjobs[0]))
                    {
	                   $objJobsFailed.childjobs[0].output | format-list | Out-File $FailedJobOutput -Force -Append
	                   $objJobsFailed.childjobs[0].warning | format-list | Out-File $FailedJobOutput -Force -Append
	                   $objJobsFailed.childjobs[0].error | format-list | Out-File $FailedJobOutput -Force -Append
			        }
                    $ErrorText = $objJobsFailed.name + "`n"
			        $ErrorText += "Job failed"
			        $ErrorLog = New-Object System.Diagnostics.EventLog('Application')
			        $ErrorLog.MachineName = "."
			        $ErrorLog.Source = "ExDC"
			        Try{$ErrorLog.WriteEntry($ErrorText,"Error", 500)} catch{}
			        Remove-Job -Id $objJobsFailed.id
		        }
	        }

			# Remove Completed Jobs
	        $colJobsCompleted = @((Get-Job -State completed) | where-object {$null -ne $_.childjobs})
	        foreach ($objJobsCompleted in $colJobsCompleted)
	        {
		        Remove-Job -Id $objJobsCompleted.id
		        write-host "---- Finished job " $objJobsCompleted.name -ForegroundColor Green
			}

			# Get Running Jobs
			$colJobsRunning = @((Get-Job -State Running) | where-object {$_.Module -ne "__DynamicModule*"})
	        if ((Test-Path ".\RunningJobs.txt") -eq $false)
	        {
		        new-item ".\RunningJobs.txt" -type file -Force
	        }
	        $RunningJobsOutput = ""
	        $Now = Get-Date
	        foreach ($objJobsRunning in $colJobsRunning)
	        {
		        $JobPID = $objJobsRunning.childjobs[0].output[0]
		        if ($null -ne $JobPID)
		        {
			        # Pass the variable assignment as a condition to reduce timing issues
			        if (($JobStartTime = ((Get-Process | where-object {$_.id -eq $JobPID}).starttime)) -ne $null)
			        {
				        $JobRunningTime = [int](($Now - $JobStartTime).TotalMinutes)
				        if ((($objJobsRunning.childjobs[0].output[1] -eq "WMI") -and ($JobRunningTime -gt ($intWMIJobTimeout/60))) `
					        -or (($objJobsRunning.childjobs[0].output[1] -eq "ExOrg") -and ($JobRunningTime -gt ($intExchJobTimeout/60))))
				        {
					        try
					        {
						        (Get-Process | where-object {$_.id -eq $JobPID}).kill()
						        write-host "Timer expired.  Killing job process $JobPID - " + $objJobsRunning.name -ForegroundColor Red
						        $ErrorText = $objJobsRunning.name + "`n"
						        $ErrorText += "Process $JobPID killed`n"
						        if ($objJobsRunning.childjobs[0].output[1] -eq "WMI") {$ErrorText += "Timeout $intWMIJobTimeout seconds exceeded"}
						        if ($objJobsRunning.childjobs[0].output[1] -eq "ExOrg") {$ErrorText += "Timeout $intExchJobTimeout seconds exceeded"}
						        $ErrorLog = New-Object System.Diagnostics.EventLog('Application')
						        $ErrorLog.MachineName = "."
						        $ErrorLog.Source = "ExDC"
						        Try{$ErrorLog.WriteEntry($ErrorText,"Error", 600)} catch{}
					        }
					        catch [System.Management.Automation.MethodInvocationException]
					        {
						        write-host "`tMethodInvocationException occured during Kill request for process $JobPID" -ForegroundColor Red
					        }
					        catch [System.Management.Automation.RuntimeException]
					        {
						        write-host "`tRuntimeException occured during Kill request for process $JobPID" -ForegroundColor Red
					        }
				        }
				        $RunningJobsOutput += "Job Name: " + $objJobsRunning.name + "`n"
				        $RunningJobsOutput += "Job State: " + $objJobsRunning.State + "`n"
				        $RunningJobsOutput += "Job process PID: " + $JobPID + "`n"
				        $RunningJobsOutput += "Job process time running: " +  $JobRunningTime + " min"
				        $RunningJobsOutput += "`n`n"
			        }
		        }
	        }
	        $RunningJobsOutput | Out-File ".\RunningJobs.txt" -Force

			# Limit Jobs based on PSSessions instead of job counts
			$intRunningJobs = Get-PsSessionCount -CurrentUser $CurrentUser -PsSession $PsSession -ResourceUri $ResourceUri
			#$intRunningJobs = $colJobsRunning.count
	        if ($intRunningJobs -eq $null)
	        {
		        $intRunningJobs = "0"
	        }
        if ($intRunningJobs -ge $JobThrottleMaxJobs)
        {
            write-host "** Throttling at $intRunningJobs jobs." -ForegroundColor DarkYellow
            Start-Sleep -Seconds $JobThrottlePolling
        }
	} while ($intRunningJobs -ge $JobThrottleMaxJobs)


	write-host "** $intRunningJobs jobs running." -ForegroundColor DarkYellow

}

Function Update-ExDCJobCount
{
	Param([int]$JobCountMaxJobs,`
		[int]$JobCountPolling)

	# This function really just checks for job completion before returning control to GUI
	# Remove Failed Jobs
	$colJobsFailed = @(Get-Job -State Failed)
	foreach ($objJobsFailed in $colJobsFailed)
	{
		if ($objJobsFailed.module -like "__DynamicModule*")
		{
			Remove-Job -Id $objJobsFailed.id
		}
		else
		{
            write-host "---- Failed job " $objJobsFailed.name -ForegroundColor Red
			$FailedJobOutput = ".\FailedJobs_" + $append + ".txt"
			if ((Test-Path $FailedJobOutput) -eq $false)
	        {
		      new-item $FailedJobOutput -type file -Force
	        }
	        "* * * * * * * * * * * * * * * * * * * *"  | Out-File $FailedJobOutput -Force -Append
	        "Job Name: " + $objJobsFailed.name | Out-File $FailedJobOutput -Force -Append
	        "Job State: " + $objJobsFailed.state | Out-File $FailedJobOutput -Force	-Append
            if ($null -ne ($objJobsFailed.childjobs[0]))
            {
	           $objJobsFailed.childjobs[0].output | format-list | Out-File $FailedJobOutput -Force -Append
	           $objJobsFailed.childjobs[0].warning | format-list | Out-File $FailedJobOutput -Force -Append
	           $objJobsFailed.childjobs[0].error | format-list | Out-File $FailedJobOutput -Force -Append
			}
            $ErrorText = $objJobsFailed.name + "`n"
			$ErrorText += "Job failed"
			$ErrorLog = New-Object System.Diagnostics.EventLog('Application')
			$ErrorLog.MachineName = "."
			$ErrorLog.Source = "ExDC"
			Try{$ErrorLog.WriteEntry($ErrorText,"Error", 500)} catch{}
			Remove-Job -Id $objJobsFailed.id
		}
	}

	$colJobsRunning = @((Get-Job -State Running) | where-object {$_.Module -ne "__DynamicModule*"})
	if ((Test-Path ".\RunningJobs.txt") -eq $false)
	{
		new-item ".\RunningJobs.txt" -type file -Force
	}

	$RunningJobsOutput = ""
	$Now = Get-Date
	foreach ($objJobsRunning in $colJobsRunning)
	{
		$JobPID = $objJobsRunning.childjobs[0].output[0]
		if ($null -ne $JobPID)
		{
			# Pass the variable assignment as a condition to reduce timing issues
			if (($JobStartTime = ((Get-Process | where-object {$_.id -eq $JobPID}).starttime)) -ne $null)
			{
				$JobRunningTime = [int](($Now - $JobStartTime).TotalMinutes)
				if ((($objJobsRunning.childjobs[0].output[1] -eq "WMI") -and ($JobRunningTime -gt ($intWMIJobTimeout/60))) `
					-or (($objJobsRunning.childjobs[0].output[1] -eq "ExOrg") -and ($JobRunningTime -gt ($intExchJobTimeout/60))))
				{
					try
                    {
                        (Get-Process | where-object {$_.id -eq $JobPID}).kill()
                    }
                    catch {}
					write-host "Timer expired.  Killing job process $JobPID - " $objJobsRunning.name -ForegroundColor Red
					$ErrorText = $objJobsRunning.name + "`n"
					$ErrorText += "Process $JobPID killed`n"
					if ($objJobsRunning.childjobs[0].output[1] -eq "WMI") {$ErrorText += "Timeout $intWMIJobTimeout seconds exceeded"}
					if ($objJobsRunning.childjobs[0].output[1] -eq "ExOrg") {$ErrorText += "Timeout $intExchJobTimeout seconds exceeded"}
					$ErrorLog = New-Object System.Diagnostics.EventLog('Application')
					$ErrorLog.MachineName = "."
					$ErrorLog.Source = "ExDC"
					Try{$ErrorLog.WriteEntry($ErrorText,"Error", 600)} catch{}
				}
				$RunningJobsOutput += "Job Name: " + $objJobsRunning.name + "`n"
				$RunningJobsOutput += "Job State: " + $objJobsRunning.State + "`n"
				$RunningJobsOutput += "Job process PID: " + $JobPID + "`n"
				$RunningJobsOutput += "Job process time running: " +  $JobRunningTime + " min"
				$RunningJobsOutput += "`n`n"
			}
		}
	}
	$RunningJobsOutput | Out-File ".\RunningJobs.txt" -Force

	$intJobCount = $colJobsRunning.count
	if ($intJobCount -eq $null)
	{
		$intJobCount = "0"
	}

	$colJobsCompleted = @((Get-Job -State completed) | where-object {$null -ne $_.childjobs})
	foreach ($objJobsCompleted in $colJobsCompleted)
	{
		Remove-Job -Id $objJobsCompleted.id
		write-host "---- Finished job " $objJobsCompleted.name -ForegroundColor Green
	}

	if ($intJobCount -ge $JobCountMaxJobs)
	{
		write-host "** $intJobCount jobs still running.  Time: $((Get-Date).timeofday.tostring())" -ForegroundColor DarkYellow
		Start-Sleep -Seconds $JobCountPolling
		Update-ExDCJobCount $JobCountMaxJobs $JobCountPolling
	}
}

Function Get-DCBoxStatus # See if any are checked
{
if (($chk_DC_Win32_Bios.checked -eq $true) -or
	($chk_DC_Win32_ComputerSystem.checked -eq $true) -or
	($chk_DC_Win32_LogicalDisk.checked -eq $true) -or
	($chk_DC_Win32_NetworkAdapter.checked -eq $true) -or
	($chk_DC_Win32_NetworkAdapterConfig.checked -eq $true) -or
	($chk_DC_Win32_OperatingSystem.checked -eq $true) -or
	($chk_DC_Win32_PageFileUsage.checked -eq $true) -or
	($chk_DC_Win32_PhysicalMemory.checked -eq $true) -or
	($chk_DC_Win32_Processor.checked -eq $true) -or
	($chk_DC_Registry_AD.checked -eq $true) -or
	($chk_DC_Registry_OS.checked -eq $true) -or
	($chk_DC_Registry_Software.checked -eq $true) -or
	($chk_DC_MicrosoftDNS_Zone.checked -eq $true) -or
    ($chk_DC_MSAD_DomainController.checked -eq $true) -or
	($chk_DC_MSAD_ReplNeighbor.checked -eq $true))
	{
		$true
	}
}

Function Get-ExBoxStatus # See if any are checked
{
if (($chk_Ex_Win32_Bios.checked -eq $true) -or
	($chk_Ex_Win32_ComputerSystem.checked -eq $true) -or
	($chk_Ex_Win32_LogicalDisk.checked -eq $true) -or
	($chk_Ex_Win32_NetworkAdapter.checked -eq $true) -or
	($chk_Ex_Win32_NetworkAdapterConfig.checked -eq $true) -or
	($chk_Ex_Win32_OperatingSystem.checked -eq $true) -or
	($chk_Ex_Win32_PageFileUsage.checked -eq $true) -or
	($chk_Ex_Win32_PhysicalMemory.checked -eq $true) -or
	($chk_Ex_Win32_Processor.checked -eq $true) -or
	($chk_Ex_Registry_Ex.checked -eq $true) -or
	($chk_Ex_Registry_OS.checked -eq $true) -or
	($chk_Ex_Registry_Software.checked -eq $true))
	{
		$true
	}
}

Function Get-ExOrgBoxStatus # See if any are checked
{
if (($chk_Org_Get_AcceptedDomain.checked -eq $true) -or
	($chk_Org_Get_ActiveSyncDevice.checked -eq $true) -or
	($chk_Org_Get_ActiveSyncPolicy.checked -eq $true) -or
	($chk_Org_Get_ActiveSyncVirtualDirectory.checked -eq $true) -or
	($chk_Org_Get_AddressBookPolicy.checked -eq $true) -or
	($chk_Org_Get_AddressList.checked -eq $true) -or
	($chk_Org_Get_AdPermission.checked -eq $true) -or
	($chk_Org_Get_AdSite.checked -eq $true) -or
	($chk_Org_Get_AdSiteLink.checked -eq $true) -or
	($chk_Org_Get_AutodiscoverVirtualDirectory.checked -eq $true) -or
	($chk_Org_Get_AvailabilityAddressSpace.checked -eq $true) -or
	($chk_Org_Get_CalendarProcessing.checked -eq $true) -or
	($chk_Org_Get_CASMailbox.checked -eq $true) -or
	($chk_Org_Get_ClientAccessArray.checked -eq $true) -or
	($chk_Org_Get_ClientAccessServer.checked -eq $true) -or
	($chk_Org_Get_ContentFilterConfig.checked -eq $true) -or
	($chk_Org_Get_DatabaseAvailabilityGroup.checked -eq $true) -or
	($chk_Org_Get_DAGNetwork.checked -eq $true) -or
	($chk_Org_Get_DistributionGroup.checked -eq $true) -or
	($chk_Org_Get_DynamicDistributionGroup.checked -eq $true) -or
	($chk_Org_Get_ECPVirtualDirectory.checked -eq $true) -or
	($chk_Org_Get_EmailAddressPolicy.checked -eq $true) -or
	($chk_Org_Get_ExchangeCertificate.checked -eq $true) -or
	($chk_Org_Get_ExchangeServer.checked -eq $true) -or
	($chk_Org_Get_Mailbox.checked -eq $true) -or
	($chk_Org_Get_MailboxDatabase.checked -eq $true) -or
	($chk_Org_Get_MailboxDatabaseCopyStatus.checked -eq $true) -or
	($chk_Org_Get_MailboxFolderStatistics.checked -eq $true) -or
	($chk_Org_Get_MailboxPermission.checked -eq $true) -or
	($chk_Org_Get_MailboxServer.checked -eq $true) -or
	($chk_Org_Get_MailboxStatistics.checked -eq $true) -or
	($chk_Org_Get_OABVirtualDirectory.checked -eq $true) -or
	($chk_Org_Get_OfflineAddressBook.checked -eq $true) -or
	($chk_Org_Get_OrgConfig.checked -eq $true) -or
	($chk_Org_Get_OutlookAnywhere.checked -eq $true) -or
	($chk_Org_Get_OwaMailboxPolicy.checked -eq $true) -or
	($chk_Org_Get_OWAVirtualDirectory.checked -eq $true) -or
	($chk_Org_Get_PowershellVirtualDirectory.checked -eq $true) -or
	($chk_Org_Get_PublicFolder.checked -eq $true) -or
	($chk_Org_Get_PublicFolderDatabase.checked -eq $true) -or
	($chk_Org_Get_PublicFolderStatistics.checked -eq $true) -or
	($chk_Org_Get_ReceiveConnector.checked -eq $true) -or
	($chk_Org_Get_RemoteDomain.checked -eq $true) -or
	($chk_Org_Get_Rbac.checked -eq $true) -or
	($chk_Org_Get_RetentionPolicy.checked -eq $true) -or
	($chk_Org_Get_RetentionPolicyTag.checked -eq $true) -or
	($chk_Org_Get_RoutingGroupConnector.checked -eq $true) -or
	($chk_Org_Get_RPCClientAccess.checked -eq $true) -or
	($chk_Org_Get_SendConnector.checked -eq $true) -or
	($chk_Org_Get_StorageGroup.checked -eq $true) -or
	($chk_Org_Get_ThrottlingPolicy.checked -eq $true) -or
	($chk_Org_Get_TransportConfig.checked -eq $true) -or
	($chk_Org_Get_TransportRule.checked -eq $true) -or
	($chk_Org_Get_TransportServer.checked -eq $true) -or
	($chk_Org_Get_UmAutoAttendant.checked -eq $true) -or
	($chk_Org_Get_UmDialPlan.checked -eq $true) -or
	($chk_Org_Get_UmIpGateway.checked -eq $true) -or
	($chk_Org_Get_UmMailbox.checked -eq $true) -or
	#($chk_Org_Get_UmMailboxConfiguration.checked -eq $true) -or
	#($chk_Org_Get_UmMailboxPin.checked -eq $true) -or
	($chk_Org_Get_UmMailboxPolicy.checked -eq $true) -or
	($chk_Org_Get_UmServer.checked -eq $true) -or
	($chk_Org_Get_User.checked -eq $true) -or
	($chk_Org_Get_WebServicesVirtualDirectory.checked -eq $true) -or
	($chk_Org_Quota.checked -eq $true) -or
	($chk_Org_Get_Fsmo.checked -eq $true) -or
	($chk_Org_Get_ExchangeServerBuilds.checked -eq $true) -or
	($chk_Org_Get_AdminGroups.checked -eq $true))	{
		$true
	}
}

Function Get-ExOrgMbxBoxStatus # See if any are checked
{
if (($chk_Org_Get_AdPermission.checked -eq $true) -or
	($chk_Org_Get_CalendarProcessing.checked -eq $true) -or
	($chk_Org_Get_CASMailbox.checked -eq $true) -or
	($chk_Org_Get_Mailbox.checked -eq $true) -or
	($chk_Org_Get_MailboxFolderStatistics.checked -eq $true) -or
	($chk_Org_Get_MailboxPermission.checked -eq $true) -or
	($chk_Org_Get_MailboxStatistics.checked -eq $true) -or
	($chk_Org_Get_UmMailbox.checked -eq $true) -or
	#($chk_Org_Get_UmMailboxConfiguration.checked -eq $true) -or
	#($chk_Org_Get_UmMailboxPin.checked -eq $true) -or
	($chk_Org_Get_User.checked -eq $true) -or
	($chk_Org_Quota.checked -eq $true))
	{
		$true
	}
}

Function Get-ExOrgServerStatus # See if any are checked
{
if (($chk_Org_Get_MailboxDatabase.checked -eq $true) -or
	($chk_Org_Get_PublicFolderDatabase.checked -eq $true) -or
	($chk_Org_Get_StorageGroup.checked -eq $true) -or
	($chk_Org_Get_ActiveSyncVirtualDirectory.checked -eq $true) -or
	($chk_Org_Get_AutodiscoverVirtualDirectory.checked -eq $true) -or
	($chk_Org_Get_OABVirtualDirectory.checked -eq $true) -or
	($chk_Org_Get_OWAVirtualDirectory.checked -eq $true) -or
	($chk_Org_Get_PublicFolder.checked -eq $true) -or
	($chk_Org_Get_PublicFolderStatistics.checked -eq $true) -or
	($chk_Org_Get_WebServicesVirtualDirectory.checked -eq $true) -or
	($chk_Org_Get_ExchangeCertificate.checked -eq $true) -or
	($chk_Org_Get_ECPVirtualDirectory.checked -eq $true) -or
	($chk_Org_Get_MailboxDatabaseCopyStatus.checked -eq $true) -or
	($chk_Org_Get_PowershellVirtualDirectory.checked -eq $true))
	{
		$true
	}
}

Function Get-ClusterBoxStatus # See if any are checked
{
if (($chk_Cluster_MSCluster_Node.checked -eq $true) -or
	($chk_Cluster_MSCluster_Network.checked -eq $true) -or
	($chk_Cluster_MSCluster_Resource.checked -eq $true) -or
	($chk_Cluster_MSCluster_ResourceGroup.checked -eq $true))
	{
		$true
	}
}

Function Import-TargetsDc
{
	Disable-AllTargetsButtons
    $status_Step1.Text = "Step 1 Status: Running"

	if ((Test-Path ".\dc.txt") -eq $true)
	{
		$EventLog = New-Object System.Diagnostics.EventLog('Application')
		$EventLog.MachineName = "."
		$EventLog.Source = "ExDC"
		try{$EventLog.WriteEntry("Starting ExDC Step 1 - Populate Domain Controllers","Information", 10)} catch{}
		$array_DC_Filtered = $null
		$File_Location = $location + "\dc.txt"
        $array_DC = @([System.IO.File]::ReadAllLines($File_Location))
		foreach ($member_DC in $array_DC)
		{
			if ($member_DC -ne "")
			{
				[array]$array_DC_Filtered += $member_DC
			}
		}
		$clb_Step1_DC_List.items.clear()
		$global:intDCTotal = $array_DC_Filtered.length
		foreach ($member_DC_Filtered in $array_DC_Filtered)
	    {
			$clb_Step1_DC_List.items.add($member_DC_Filtered)
		}
		For ($i=0;$i -le ($intDCTotal - 1);$i++)
		{
			$clb_Step1_DC_List.SetItemChecked($i,$true)
		}
		$EventLog = New-Object System.Diagnostics.EventLog('Application')
		$EventLog.MachineName = "."
		$EventLog.Source = "ExDC"
		try{$EventLog.WriteEntry("Ending ExDC Step 1 - Populate Domain Controllers","Information", 11)} catch{}
		$txt_DCTotal.Text = "Domain Controller count = " + $intDCTotal
		$txt_DCTotal.visible = $true
	    $status_Step1.Text = "Step 1 Status: Idle"
	}
	else
	{
		write-host	"The file dc.txt is not present.  Run Discover to create the file."
		$status_Step1.Text = "Step 1 Status: Failed - dc.txt file not found.  Run Discover to create the file."
	}
	Enable-AllTargetsButtons
}

Function Import-TargetsEx
{
	Disable-AllTargetsButtons
    $status_Step1.Text = "Step 1 Status: Running"

	if ((Test-Path ".\exchange.txt") -eq $true)
	{
		$EventLog = New-Object System.Diagnostics.EventLog('Application')
		$EventLog.MachineName = "."
		$EventLog.Source = "ExDC"
		try{$EventLog.WriteEntry("Starting ExDC Step 1 - Populate Exchange servers","Information", 10)} catch{}

		$array_Ex_Filtered = $null
        $File_Location = $location + "\exchange.txt"
		$array_Ex = @([System.IO.File]::ReadAllLines($File_Location))
		foreach ($member_Ex in $array_Ex)
		{
			if ($member_Ex -ne "")
			{
				[array]$array_Ex_Filtered += $member_Ex
			}
		}
		$clb_Step1_Ex_List.items.clear()
		$global:intExTotal = $array_Ex_Filtered.length
		foreach ($member_Ex_Filtered in $array_Ex_Filtered)
	    {
			$clb_Step1_Ex_List.items.add($member_Ex_Filtered)
		}
		For ($i=0;$i -le ($intExTotal - 1);$i++)
		{
			$clb_Step1_Ex_List.SetItemChecked($i,$true)
		}
		$EventLog = New-Object System.Diagnostics.EventLog('Application')
		$EventLog.MachineName = "."
		$EventLog.Source = "ExDC"
		try{$EventLog.WriteEntry("Ending ExDC Step 1 - Populate Exchange servers","Information", 11)} catch{}

		$txt_ExchTotal.Text = "Exchange server count = " + $intExTotal
		$txt_ExchTotal.visible = $true
	    $status_Step1.Text = "Step 1 Status: Idle"
	}
	else
	{
		write-host	"The file exchange.txt is not present.  Run Discover to create the file."
		$status_Step1.Text = "Step 1 Status: Failed - exchange.txt file not found.  Run Discover to create the file."
	}
	Enable-AllTargetsButtons
}

Function Import-TargetsNodes
{
	Disable-AllTargetsButtons
    $status_Step1.Text = "Step 1 Status: Running"
	if ((Test-Path ".\ClusterNodes.txt") -eq $true)
	{
		$EventLog = New-Object System.Diagnostics.EventLog('Application')
		$EventLog.MachineName = "."
		$EventLog.Source = "ExDC"
		try{$EventLog.WriteEntry("Starting ExDC Step 1 - Populate Exchange nodes","Information", 10)} catch{}
		$array_Nodes_Filtered = $null
        $File_Location = $location + "\ClusterNodes.txt"
		$array_Nodes = @([System.IO.File]::ReadAllLines($File_Location))
		foreach ($member_Nodes in $array_Nodes)
		{
			if ($member_Nodes -ne "")
			{
				[array]$array_Nodes_Filtered += $member_Nodes
			}
		}
		$clb_Step1_Nodes_List.items.clear()
		$global:intNodesTotal = $array_Nodes_Filtered.length
		if ($intNodesTotal -gt 0)
		{
			foreach ($member_Nodes_Filtered in $array_Nodes_Filtered)
	    	{
				$clb_Step1_Nodes_List.items.add($member_Nodes_Filtered)
			}
			For ($i=0;$i -le ($intNodesTotal - 1);$i++)
			{
				$clb_Step1_Nodes_List.SetItemChecked($i,$true)
			}
		}
		else
		{
			$intNodesTotal = 0
		}
		$EventLog = New-Object System.Diagnostics.EventLog('Application')
		$EventLog.MachineName = "."
		$EventLog.Source = "ExDC"
		try{$EventLog.WriteEntry("Ending ExDC Step 1 - Populate Exchange nodes","Information", 11)} catch{}

		$txt_NodesTotal.Text = "Exchange node count = " + $intNodesTotal
		$txt_NodesTotal.visible = $true
	    $status_Step1.Text = "Step 1 Status: Idle"
	}
	else
	{
		write-host	"The file ClusterNodes.txt is not present.  Run Discover to create the file."
		$status_Step1.Text = "Step 1 Status: Failed - ClusterNodes.txt file not found.  Run Discover to create the file."
	}
	Enable-AllTargetsButtons
}

Function Import-TargetsMailboxes
{
	Disable-AllTargetsButtons
    $status_Step1.Text = "Step 1 Status: Running"
	$File_Location = $location + "\mailbox.txt"
    if ((Test-Path $File_Location) -eq $true)
	{
		$EventLog = New-Object System.Diagnostics.EventLog('Application')
		$EventLog.MachineName = "."
		$EventLog.Source = "ExDC"
		try{$EventLog.WriteEntry("Starting ExDC Step 1 - Populate","Information", 10)} catch{}
	    $array_Mailboxes = @(([System.IO.File]::ReadAllLines($File_Location)) | sort-object -Unique)
		$global:intMailboxTotal = 0
	    $clb_Step1_Mailboxes_List.items.clear()
		foreach ($member_Mailbox in $array_Mailboxes | where-object {$_ -ne ""})
	    {
	        $clb_Step1_Mailboxes_List.items.add($member_Mailbox)
			$global:intMailboxTotal++
	    }
		For ($i=0;$i -le ($intMailboxTotal - 1);$i++)
		{
			$clb_Step1_Mailboxes_List.SetItemChecked($i,$true)
		}
		$EventLog = New-Object System.Diagnostics.EventLog('Application')
		$EventLog.MachineName = "."
		$EventLog.Source = "ExDC"
		try{$EventLog.WriteEntry("Ending ExDC Step 1 - Populate","Information", 11)} catch{}
		$txt_MailboxesTotal.Text = "Mailbox count = " + $intMailboxTotal
		$txt_MailboxesTotal.visible = $true
	    $status_Step1.Text = "Step 2 Status: Idle"
	}
	else
	{
		write-host	"The file mailbox.txt is not present.  Run Discover to create the file."
		$status_Step1.Text = "Step 1 Status: Failed - mailbox.txt file not found.  Run Discover to create the file."
	}
	Enable-AllTargetsButtons
}


Function Enable-TargetsDc
{
	For ($i=0;$i -le ($intDCTotal -1);$i++)
	{
		$clb_Step1_DC_List.SetItemChecked($i,$true)
	}
}

Function Disable-TargetsDc
{
	For ($i=0;$i -le ($intDCTotal -1);$i++)
	{
		$clb_Step1_DC_List.SetItemChecked($i,$False)
	}
}

Function Enable-TargetsEx
{
	For ($i=0;$i -le ($intExTotal -1);$i++)
	{
		$clb_Step1_Ex_List.SetItemChecked($i,$true)
	}
}

Function Disable-TargetsEx
{
	For ($i=0;$i -le ($intExTotal -1);$i++)
	{
		$clb_Step1_Ex_List.SetItemChecked($i,$False)
	}
}

Function Enable-TargetsNodes
{
	For ($i=0;$i -le ($intNodesTotal -1);$i++)
	{
		$clb_Step1_Nodes_List.SetItemChecked($i,$true)
	}
}

Function Disable-TargetsNodes
{
	For ($i=0;$i -le ($intNodesTotal -1);$i++)
	{
		$clb_Step1_Nodes_List.SetItemChecked($i,$False)
	}
}

Function Enable-TargetsMailbox
{
	For ($i=0;$i -le ($intMailboxTotal - 1);$i++)
	{
		$clb_Step1_Mailboxes_List.SetItemChecked($i,$true)
	}
}

Function Disable-TargetsMailbox
{
	For ($i=0;$i -le ($intMailboxTotal - 1);$i++)
	{
		$clb_Step1_Mailboxes_List.SetItemChecked($i,$False)
	}
}

Function Set-AllFunctionsDc
{
    Param([boolean]$Check)
	$chk_DC_Win32_Bios.checked = $Check
	$chk_DC_Win32_ComputerSystem.checked = $Check
	$chk_DC_Win32_LogicalDisk.checked = $Check
	$chk_DC_Win32_NetworkAdapter.checked = $Check
	$chk_DC_Win32_NetworkAdapterConfig.checked = $Check
	$chk_DC_Win32_OperatingSystem.checked = $Check
	$chk_DC_Win32_PageFileUsage.checked = $Check
	$chk_DC_Win32_PhysicalMemory.checked = $Check
	$chk_DC_Win32_Processor.checked = $Check
	$chk_DC_Registry_AD.checked = $Check
	$chk_DC_Registry_OS.checked = $Check
	$chk_DC_Registry_Software.checked = $Check
	$chk_DC_MicrosoftDNS_Zone.checked = $Check
	$chk_DC_MSAD_DomainController.checked = $Check
	$chk_DC_MSAD_ReplNeighbor.checked = $Check
}

Function Set-AllFunctionsEx
{
    Param([boolean]$Check)
	$chk_Ex_Win32_Bios.checked = $Check
	$chk_Ex_Win32_ComputerSystem.checked = $Check
	$chk_Ex_Win32_LogicalDisk.checked = $Check
	$chk_Ex_Win32_NetworkAdapter.checked = $Check
	$chk_Ex_Win32_NetworkAdapterConfig.checked = $Check
	$chk_Ex_Win32_OperatingSystem.checked = $Check
	$chk_Ex_Win32_PageFileUsage.checked = $Check
	$chk_Ex_Win32_PhysicalMemory.checked = $Check
	$chk_Ex_Win32_Processor.checked = $Check
	$chk_Ex_Registry_Ex.checked = $Check
	$chk_Ex_Registry_OS.checked = $Check
	$chk_Ex_Registry_Software.checked = $Check
}

Function Set-AllFunctionsCluster
{
    Param([boolean]$Check)
	$chk_Cluster_MSCluster_Node.checked = $Check
	$chk_Cluster_MSCluster_Network.checked = $Check
	$chk_Cluster_MSCluster_Resource.checked = $Check
	$chk_Cluster_MSCluster_ResourceGroup.checked = $Check
}

Function Set-AllFunctionsClientAccess
{
    Param([boolean]$Check)
	$chk_Org_Get_ActiveSyncDevice.Checked = $Check
	$chk_Org_Get_ActiveSyncPolicy.Checked = $Check
	$chk_Org_Get_ActiveSyncVirtualDirectory.Checked = $Check
	$chk_Org_Get_AutodiscoverVirtualDirectory.Checked = $Check
	$chk_Org_Get_AvailabilityAddressSpace.Checked = $Check
	$chk_Org_Get_ClientAccessArray.Checked = $Check
	$chk_Org_Get_ClientAccessServer.Checked = $Check
	$chk_Org_Get_ECPVirtualDirectory.Checked = $Check
	$chk_Org_Get_OABVirtualDirectory.Checked = $Check
	$chk_Org_Get_OutlookAnywhere.Checked = $Check
	$chk_Org_Get_OwaMailboxPolicy.Checked = $Check
	$chk_Org_Get_OWAVirtualDirectory.Checked = $Check
	$chk_Org_Get_PowershellVirtualDirectory.Checked = $Check
	$chk_Org_Get_RPCClientAccess.Checked = $Check
	$chk_Org_Get_ThrottlingPolicy.Checked = $Check
	$chk_Org_Get_WebServicesVirtualDirectory.Checked = $Check
}

Function Set-AllFunctionsGlobal
{
    Param([boolean]$Check)
	$chk_Org_Get_AddressBookPolicy.Checked = $Check
	$chk_Org_Get_AddressList.Checked = $Check
	$chk_Org_Get_DatabaseAvailabilityGroup.Checked = $Check
	$chk_Org_Get_DAGNetwork.Checked = $Check
	$chk_Org_Get_EmailAddressPolicy.Checked = $Check
	$chk_Org_Get_ExchangeCertificate.Checked = $Check
	$chk_Org_Get_ExchangeServer.Checked = $Check
	$chk_Org_Get_MailboxDatabase.Checked = $Check
	$chk_Org_Get_MailboxDatabaseCopyStatus.Checked = $Check
	$chk_Org_Get_MailboxServer.Checked = $Check
	$chk_Org_Get_OfflineAddressBook.Checked = $Check
	$chk_Org_Get_OrgConfig.Checked = $Check
	$chk_Org_Get_PublicFolderDatabase.Checked = $Check
	$chk_Org_Get_Rbac.Checked = $Check
	$chk_Org_Get_RetentionPolicy.Checked = $Check
	$chk_Org_Get_RetentionPolicyTag.Checked = $Check
	$chk_Org_Get_StorageGroup.Checked = $Check
}

Function Set-AllFunctionsRecipient
{
    Param([boolean]$Check)
	$chk_Org_Get_ADPermission.Checked = $Check
	$chk_Org_Get_CalendarProcessing.Checked = $Check
	$chk_Org_Get_CASMailbox.Checked = $Check
	$chk_Org_Get_DistributionGroup.Checked = $Check
	$chk_Org_Get_DynamicDistributionGroup.Checked = $Check
	$chk_Org_Get_Mailbox.Checked = $Check
	$chk_Org_Get_MailboxFolderStatistics.Checked = $Check
	$chk_Org_Get_MailboxPermission.Checked = $Check
	$chk_Org_Get_MailboxStatistics.Checked = $Check
	$chk_Org_Get_PublicFolder.Checked = $Check
	$chk_Org_Get_PublicFolderStatistics.Checked = $Check
	$chk_Org_Get_User.Checked = $Check
	$chk_Org_Quota.Checked = $Check
}

Function Set-AllFunctionsTransport
{
    Param([boolean]$Check)
	$chk_Org_Get_AcceptedDomain.Checked = $Check
	$chk_Org_Get_AdSite.Checked = $Check
	$chk_Org_Get_AdSiteLink.Checked = $Check
	$chk_Org_Get_ContentFilterConfig.Checked = $Check
	$chk_Org_Get_ReceiveConnector.Checked = $Check
	$chk_Org_Get_RemoteDomain.Checked = $Check
	$chk_Org_Get_RoutingGroupConnector.Checked = $Check
	$chk_Org_Get_SendConnector.Checked = $Check
	$chk_Org_Get_TransportConfig.Checked = $Check
	$chk_Org_Get_TransportRule.Checked = $Check
	$chk_Org_Get_TransportServer.Checked = $Check
}

Function Set-AllFunctionsUm
{
    Param([boolean]$Check)
	$chk_Org_Get_UmAutoAttendant.Checked = $Check
	$chk_Org_Get_UmDialPlan.Checked = $Check
	$chk_Org_Get_UmIpGateway.Checked = $Check
	$chk_Org_Get_UmMailbox.Checked = $Check
	#$chk_Org_Get_UmMailboxConfiguration.Checked = $Check
	#$chk_Org_Get_UmMailboxPin.Checked = $Check
	$chk_Org_Get_UmMailboxPolicy.Checked = $Check
	$chk_Org_Get_UmServer.Checked = $Check
}

Function Set-AllFunctionsMisc
{
    Param([boolean]$Check)
	$chk_Org_Get_AdminGroups.Checked = $Check
	$chk_Org_Get_Fsmo.Checked = $Check
	$chk_Org_Get_ExchangeServerBuilds.Checked = $Check
}

Function Start-ExDCJob
{
    param(  [string]$server,`
            [string]$Job,`              # e.g. "Win32_ComputerSystem"
            [boolean]$JobType,`         # 0=WMI, 1=ExOrg
            [string]$Location,`
            [string]$JobScriptName,`    # e.g. "dc_w32_cs.ps1"
            [int]$i,`                   # Number or $null
            [string]$PSSession)

	#Start-sleep 3
	If ($JobType -eq 0) #WMI
        {Limit-ExDCJob -JobThrottleMaxJobs $intWMIJobs -JobThrottlePolling $intWMIPolling -PsSession $null}
    else                #ExOrg
        {Limit-ExDCJob -JobThrottleMaxJobs $intExOrgJobs -JobThrottlePolling $intExOrgPolling -PsSession $PsSession}
    $strJobName = "$Job job for $server"
    write-host "-- Starting " $strJobName
    $PS_Loc = "$location\ExDC_Scripts\$JobScriptName"
    Start-Job -ScriptBlock {param($a,$b,$c,$d,$e) Powershell.exe -NoProfile -file $a $b $c $d $e} -ArgumentList @($PS_Loc,$location,$server,$i,$PSSession) -Name $strJobName
    start-sleep 5 # Allow time for child job to spawn
}

Function Get-PsSessionCount
{
	param(	[string]$CurrentUser,`
			[string]$PsSession,`
			[string]$ResourceUri`
			)
	# This is function is only called during Limit-ExDCCount to prevent new jobs from spawning prematurely 
	Try
	{
		$CxnUri = "http://" + $PsSession + "/powershell"
		$Sessions = (Get-WSManInstance -ConnectionURI $CxnUri shell -Enumerate) | where-object {($_.ResourceUri -eq $ResourceUri) -and ($_.owner -eq $CurrentUser)}
		return $sessions.count
	}
	catch [system.Management.Automation.ParameterBindingException]
	{
		# Fall back to job count for WMI
		Return $colJobsRunning.count
	}
}


#endregion *** Custom Functions ***

# Check Powershell version
$PowershellVersionNumber = $null
$powershellVersion = get-host
# Teminate if Powershell is less than version 2
if ($powershellVersion.Version.Major -lt "2")
{
    write-host "Unsupported Powershell version detected."
    write-host "Powershell v2 is required."
    end
}
# Powershell v2 or later required for Ex2010 environments
elseif ($powershellVersion.Version.Major -lt "3")
{
    $PowershellVersionNumber = 2
	write-host "Powershell version 2 detected" -ForegroundColor Green
}
# Powershell v3 or later required for Ex2013 or later environments
elseif ($powershellVersion.Version.Major -lt "4")
{
    $PowershellVersionNumber = 3
    write-host "Powershell version 3 detected" -ForegroundColor Green
}
elseif ($powershellVersion.Version.Major -lt "5")
{
    $PowershellVersionNumber = 4
    write-host "Powershell version 4 detected" -ForegroundColor Green
}
elseif ($powershellVersion.Version.Major -lt "6")
{
    $PowershellVersionNumber = 5
    write-host "Powershell version 5 detected" -ForegroundColor Green
}

# Check for presence of Powershell Profile and warn if present
if ((test-path $PROFILE) -eq $true)
{
	write-host "WARNING: Powershell profile detected." -ForegroundColor Red
	write-host "WARNING: All jobs will be executed using the -NoProfile switch" -ForegroundColor Red
}
else
{
	write-host "No Powershell profile detected." -ForegroundColor Green
}

# Check for the presence of the Exchange Snap-ins
if ($NoGUI -ne $true)
{
	$Exchange2007Powershell = $false
	$Exchange2010Powershell = $false
	$RegisteredSnapins = Get-PSSnapin -Registered
	foreach ($Snapin in $RegisteredSnapins)
	{
		if ($Snapin.name -like "Microsoft.Exchange.Management.PowerShell.Admin")
		{
			$Exchange2007Powershell = $true
			write-host "Exchange 2007 snap-in detected." -ForegroundColor Green
		}
		if ($Snapin.name -like "Microsoft.Exchange.Management.PowerShell.E2010")
		{
			$Exchange2010Powershell = $true
			write-host "Exchange 2010 or later snap-in detected." -ForegroundColor Green
		}
	}
	#if (($Exchange2007Powershell -eq $true) -and ($Exchange2010Powershell -eq $true))
	#{#
	#	$Exchange_2007_2010_Mixed = $true
	#}
	if (($Exchange2007Powershell -eq $false) -and ($Exchange2010Powershell -eq $false))
	{
		write-host "No Exchange snap-ins detected.  Checking active PSSessions..." -ForegroundColor Green
		If ($null -ne (get-pssession | where-object {$_.state -eq "opened" -and $_.ConfigurationName -eq "Microsoft.Exchange"}))
		{
			write-host "Open PSSession to Microsoft.Exchange detected." -foregroundcolor green
			$Exchange2010Powershell = $true
		}
		else
		{
			write-host	"No open PSSession to Microsoft.Exchange detected.  Setting -NoEMS switch" -foregroundcolor yellow
			$NoEMS = $true
		}
	}
}

#----------------------------------------------
# Write to Event Log
#----------------------------------------------
$EventLog = New-Object System.Diagnostics.EventLog('Application')
$EventLog.MachineName = "."
$EventLog.Source = "ExDC"
try{$EventLog.WriteEntry("Starting ExDC Run","Information", 1)} catch{}

if ($Exchange2010Powershell -eq $true)
{
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "ExDC"
	try{$EventLog.WriteEntry("Exchange 2010 or later snap-in detected on server","Information", 2)} catch{}
    Set-AdServerSettings -ViewEntireForest $true
}
elseif ($Exchange2007Powershell -eq $true)
{
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "ExDC"
	try{$EventLog.WriteEntry("Exchange 2007 snap-in detected on server","Information", 2)} catch{}
	([Microsoft.Exchange.Data.Directory.AdminSessionADSettings]::Instance).ViewEntireForest = $true
}

#----------------------------------------------
# Initialize Arrays and Variables
#----------------------------------------------
Set-Variable -name intExOrgJobs -Scope global
Set-Variable -name intExOrgPolling -Scope global
Set-Variable -name intJobs -Scope global
Set-Variable -name intPolling -Scope global
Set-Variable -name intWMIJobTimeout -Scope global
Set-Variable -name intExchJobTimeout -Scope global
Set-Variable -name INI -Scope global
Set-Variable -name intDCTotal -Scope global
Set-Variable -name intExTotal -Scope global
Set-Variable -name intNodesTotal -Scope global
Set-Variable -name intMailboxTotal -Scope global

$array_DC = @()
#$array_Exch = @()
#$array_ExchNodes = @()
$array_Mailboxes = @()
$Exchange2003orEarlier = $false
$Exchange2007 = $false
$Exchange2010 = $false
$Exchange2013 = $false
$Exchange2016 = $false
$UM=$true
$CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$resourceuri = "http://schemas.microsoft.com/powershell/Microsoft.Exchange"


if ($JobCount_ExOrg -eq 0) 			{$intExOrgJobs = 10}
	else 							{$intExOrgJobs = $JobCount_ExOrg}
if ($JobPolling_ExOrg -eq 0) 		{$intExOrgPolling = 5}
	else 							{$intExOrgPolling = $JobPolling_ExOrg}
if ($JobCount_WMI -eq 0) 			{$intWMIJobs = 25}
	else 							{$intWMIJobs = $JobCount_WMI}
if ($JobPolling_WMI -eq 0)			{$intWMIPolling = 5}
	else 							{$intWMIPolling = $JobPolling_WMI}
if ($Timeout_WMI_Job -eq 0) 		{$intWMIJobTimeout = 600} 				# 600 sec = 10 min
	else 							{$intWMIJobTimeout = $Timeout_WMI_Job}
	#$intWMIJobTimeoutMS = $intWMIJobTimeout * 1000 							# convert sec to ms
if ($Timeout_ExOrg_Job -eq 0)		{$intExchJobTimeout = 3600} 			# 3600 sec = 60 min
	else 							{$intExchJobTimeout = $Timeout_ExOrg_Job}
	#$intExchJobTimeoutMS = $intExchJobTimeout * 1000 						# convert sec to ms

If (($NoGUI -eq $true) -and ($ServerforPSSession -ne $null))
{
	write-host "NoGUI is specified.  Setting up the Remote Powershell session to $ServerforPSSession"
	$cxnUri = "http://" + $ServerforPSSession + "/powershell"
	$session = New-PSSession -configurationName Microsoft.Exchange -ConnectionUri $cxnUri -authentication Kerberos
	Import-PSSession -Session $session -AllowClobber
    Set-AdServerSettings -ViewEntireForest $true

}
elseif (($NoGUI -eq $true) -and ($ServerforPSSession -eq $null))
{
	write-host "ServerForPSSession must be specified when using the -NoGUI switch" -foregroundcolor red
	Exit
}

If ($NoEMS -eq $false)
{
	$EventText = ""
	write-host "`tChecking Exchange versions..." -foregroundcolor cyan
	foreach ($server in (Get-ExchangeServer))
	{
		#write-host "`tTesting $server..."
	    if ($server.IsExchange2007OrLater -eq $false)
		{$Exchange2003orEarlier = $true}
		if ((($server.IsExchange2007OrLater -eq $true) -and ($server.IsE14OrLater -eq $null)) `
			-or `
			(($server.IsExchange2007OrLater -eq $true) -and ($server.IsE14OrLater -eq $false)))
			{
	            $Exchange2007 = $true
	            #write-host "`t$server is Ex2007" -ForegroundColor Green
	        }
		if ((($server.IsExchange2007OrLater -eq $true) `
			-and ($server.IsE14OrLater -eq $true)) `
			-and ($server.AdminDisplayVersion.tostring() -match "Version 14"))
			{
	            $Exchange2010 = $true
	            #write-host "`t$server is Ex2010" -ForegroundColor Green
	        }
		if (($server.IsE15OrLater -eq $true) `
			-and ($server.AdminDisplayVersion.tostring() -match "Version 15.0"))
		{
	            $Exchange2013 = $true
	            #write-host "`t$server is Ex2013" -ForegroundColor Green
	        }
		if (($server.IsE15OrLater -eq $true) `
			-and ($server.AdminDisplayVersion.tostring() -match "Version 15.1"))
			{
	            $Exchange2016 = $true
	            #write-host "`t$server is Ex2016" -ForegroundColor Green
	        }
	}
	if ($Exchange2003orEarlier -eq $true)
	{
		write-host "Exchange 2003 or earlier detected in the environment" -ForegroundColor Cyan
		$EventText += "`tExchange 2003 or earlier detected in the environment`n"
	}
	if ($Exchange2007 -eq $true)
	{
		write-host "Exchange 2007 detected in the environment" -ForegroundColor Cyan
		$EventText += "`tExchange 2007 detected in the environment`n"
		if ($Exchange2007Powershell -eq $false)
		{
			write-host "Exchange 2007 detected but Exchange 2007 snap-in not present." -ForegroundColor Red
			write-host "Data collection may yield incorrect or incomplete results." -ForegroundColor Red
			write-host "Exchange 2010 cmdlets will still attempt to run, but fail." -ForegroundColor Red
			write-host "Run on server with both Exchange 2007 and 2010 Management Tools." -ForegroundColor Red
			$EventText += "`tExchange 2007 detected in the environment but Exchange 2007 snap-in not present.`n"
		}
	}
	if ($Exchange2010 -eq $true)
	{
		write-host "Exchange 2010 detected in the environment" -ForegroundColor Cyan
		$EventText += "`tExchange 2010 detected in the environment`n"
		if ($Exchange2010Powershell -eq $false)
		{
			write-host "Exchange 2010 detected but Exchange 2010 snap-in not present." -ForegroundColor Red
			write-host "Data collection may yield incorrect or incomplete results." -ForegroundColor Red
			write-host "Run on server with both Exchange 2007 and 2010 Management Tools." -ForegroundColor Red
			$EventText += "`tExchange 2010 Detected in the Environment but Exchange 2010 snap-in not present.`n"
		}
	}
	if (($Exchange2013 -eq $true) -or ($Exchange2016 -eq $true))
	{
		write-host "Exchange 2013 or later detected in the environment" -ForegroundColor Cyan
		$EventText += "`tExchange 2013 or later detected in the environment`n"
		if ($PowershellVersionNumber -lt 3)
		{
			write-host "Exchange 2013 or later detected but Powershell v3 or later not present." -ForegroundColor Red
			write-host "Data collection may yield incorrect or incomplete results." -ForegroundColor Red
			write-host "Please rerun on data collection server with Powershell v3 or later" -ForegroundColor Red
			$EventText += "`tExchange 2013 or later detected in the environment but Powershell v3 or later not present.`n"
		}
	}
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "ExDC"
	try{$EventLog.WriteEntry($EventText,"Information", 3)} catch{}
}
If ((($NoEMS -eq $false) -and ($Exchange2010Powershell -eq $true)) -or ($NoGUI -eq $true))
{
	if ($ServerForPSSession -ne "")	{$session_0 = $ServerForPSSession}
	else
	{
		$session = @(Get-PSSession | where-object {$_.configurationname -eq "Microsoft.Exchange"})
		$session_0 = $session[0].computername
		if (($session_0 -eq $null) -and ($Exchange2007Powershell -eq $true))
		{
			write-host "Exchange 2010 Management Shell detected but no Microsoft.Exchange configuration or Exchange snap-in could be found in this PSSession." -ForegroundColor Red
			write-host "Both Exchange 2007 and Exchange 2010 Management tools are installed on this server, please relaunch ExDC from the Exchange 2010 shell." -ForegroundColor Red
			write-host "Otherwise, please rerun and specify an Exchange server using the -ServerForPSSession switch" -ForegroundColor Red
			exit
		}
		elseif ($session_0 -eq $null)
		{
			write-host "Exchange 2010 Management Shell detected but no Microsoft.Exchange configuration or Exchange snap-in could be found in this PSSession." -ForegroundColor Red
			write-host "Please rerun and specify an Exchange server using the -ServerForPSSession switch" -ForegroundColor Red
			exit
		}
	}
}
else
{
	$session_0 = $null
}

#Set timestamp
$StartTime = Get-Date -UFormat %s
$append = $StartTime
$append = "v4_0_2." + $append
#----------------------------------------------
# Misc Code
#----------------------------------------------
$ScriptLoc = Split-Path -parent $MyInvocation.MyCommand.Definition
Set-Location $ScriptLoc
$location = [string]((get-Location).path)
$testfolder = test-path output
if ($testfolder -eq $false)
{
	new-item -name "output" -type directory -force | Out-Null
}
#Call the Function
write-host "Starting Exchange Data Collector (ExDC) v4 with the following parameters: " -ForegroundColor Cyan
$EventText = "Starting Exchange Data Collector (ExDC) v4 with the following parameters: `n"
if ($NoEMS -eq $false)
{
	write-host "`tIni Settings`t" -ForegroundColor Cyan
	$EventText += "`tIni Settings:`t" + $INI + "`n"
	write-host "`t`tServer Ini:`t" $INI_Server -ForegroundColor Cyan
	$EventText += "`t`tServer Ini:`t" + $INI_Server + "`n"
	write-host "`t`tCluster Ini:`t" $INI_Cluster -ForegroundColor Cyan
	$EventText += "`t`tCluster Ini:`t" + $INI_Cluster + "`n"
	write-host "`t`tExOrg Ini:`t" $INI_ExOrg -ForegroundColor Cyan
	$EventText += "`t`tExOrg Ini:`t" + $INI_ExOrg + "`n"
	write-host "`tNon-Exchange cmdlet jobs" -ForegroundColor Cyan
	$EventText += "`tNon-Exchange cmdlet jobs`n"
	write-host "`t`tMax jobs:`t" $intWMIJobs -ForegroundColor Cyan
	$EventText += "`t`tMax jobs:`t" + $intWMIJobs + "`n"
	write-host "`t`tPolling:`t" $intWMIPolling " seconds" -ForegroundColor Cyan
	$EventText += "`t`tPolling:`t`t" + $intWMIPolling + "`n"
	write-host "`t`tTimeout:`t" $intWMIJobTimeout " seconds" -ForegroundColor Cyan
	$EventText += "`t`tTimeout:`t" + $intWMIJobTimeout + "`n"
	write-host "`tExchange cmdlet jobs" -ForegroundColor Cyan
	$EventText += "`tExchange cmdlet jobs`n"
	write-host "`t`tMax jobs:`t" $intExOrgJobs -ForegroundColor Cyan
	$EventText += "`t`tMax jobs:`t" + $intExOrgJobs + "`n"
	write-host "`t`tPolling: `t" $intExOrgPolling " seconds" -ForegroundColor Cyan
	$EventText += "`t`tPolling: `t`t" + $intExOrgPolling + "`n"
	write-host "`t`tTimeout:`t" $intExchJobTimeout " seconds" -ForegroundColor Cyan
	$EventText += "`t`tTimeout:`t" + $intExchJobTimeout + "`n"
	if ($Exchange2003orEarlier -eq $true)
	{
		$EventText += "`tExchange 2003 or Earlier Detected in the Environment`n"
	}
	if ($Exchange2007Powershell -eq $true)
	{
		write-host "`tExchange 2007 Powershell Detected" -ForegroundColor Cyan
		$EventText += "`tExchange 2007 Powershell Detected`n"
	}
	if ($Exchange2010Powershell -eq $true)
	{
		write-host "`tExchange 2010 Powershell Detected" -ForegroundColor Cyan
		$EventText += "`tExchange 2010 Powershell Detected`n"
		write-host "`tExchange Server to use for remoting: " -ForegroundColor Cyan
		$EventText += "`tExchange Server to use for remoting:`n"

		if ($ServerForPSSession -eq "")
		{
			write-host "`t`tAutomatically Discovered" -ForegroundColor Cyan
			$EventText += "`t`tAutomatically Discovered`n"
		}
		else
		{
			write-host "`t`tManually Configured" -ForegroundColor Cyan
			$EventText += "`t`tManually Configured`n"
		}
		write-host "`t`tServer:	`t" $session_0 -ForegroundColor Cyan
		$EventText += "`t`tServer: " + $session_0
		$InitialSessionCount = Get-PsSessionCount -CurrentUser $CurrentUser -PsSession $Session_0 -ResourceUri $ResourceUri
		write-host "`tInitial PSSession count to server for this user: " $InitialSessionCount  -ForegroundColor Cyan

		# Check User Throttle Policy
		if ((get-throttlingpolicyassociation $currentuser).ThrottlingPolicyId -ne $null)
		{
			$PowershellMaxConcurrency = (Get-throttlingpolicy ((get-throttlingpolicyassociation $currentuser).ThrottlingPolicyId.Name)).PowerShellMaxConcurrency.value
		}
		else
		{
			$PowershellMaxConcurrency = (Get-ThrottlingPolicy | where-object {$_.isdefault -eq $true}).PowerShellMaxConcurrency.value
			If ($PowershellMaxConcurrency -eq $null)
			{
				$PowershellMaxConcurrency = (Get-ThrottlingPolicy -identity "GlobalThrottlingPolicy*").PowerShellMaxConcurrency.value
			}
		}
		write-host "`tPowershellMaxConcurrency for this user:`t " $PowershellMaxConcurrency  -ForegroundColor Cyan
		if ($PowershellMaxConcurrency -le ($JobCount_ExOrg + $InitialSessionCount))
		{
			write-host "`tWarning!  Job Count likely to exceed PowershellMaxConcurrency!" -foregroundcolor red
		}

	}
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "ExDC"
	try{$EventLog.WriteEntry($EventText,"Information", 4)} catch{}
}
else
{
	write-host "`tNoEMS switch used" -ForegroundColor Cyan
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "ExDC"
	try{$EventLog.WriteEntry("NoEMS switch used.","Information", 4)} catch{}
}

if ($NoGUI -eq $true)
{
	If ($ServerForPSSession -eq "")
	{
		Write-Host "NoGUI switch requires -ServerForPSSession to be set" -ForegroundColor Red
		exit
	}
	If (($INI_Server -eq "") -and ($INI_Cluster -eq "") -and ($INI_ExOrg -eq ""))
	{
		Write-Host "NoGUI switch requires at least one INI switch to be set" -ForegroundColor Red
		Exit
	}
}

# Let's start the party
New-ExDCForm