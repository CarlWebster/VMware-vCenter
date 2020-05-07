#Requires -Version 3.0
#This File is in Unicode format.  Do not edit in an ASCII editor.

#region help text

<#

.SYNOPSIS
	Creates a complete inventory of a VMware vSphere datacenter using PowerCLI and Microsoft Word 2010, 2013 or 2016.
.DESCRIPTION
	Creates a complete inventory of a VMware vSphere datacenter using PowerCLI and Microsoft Word and PowerShell.
	Creates a Word document named after the vCenter server.
	Document includes a Cover Page, Table of Contents and Footer.
	Includes support for the following language versions of Microsoft Word:
		Catalan
		Chinese
		Danish
		Dutch
		English
		Finnish
		French
		German
		Norwegian
		Portuguese
		Spanish
		Swedish
		
.PARAMETER VIServerName
    Name of the vCenter Server to connect to.
    This parameter is mandatory and does not have a default value.
    FQDN should be used; hostname can be used if it can be resolved correctly.
.PARAMETER CompanyName
	Company Name to use for the Cover Page.  
	Default value is contained in HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated on the 
	computer running the script.
	This parameter has an alias of CN.
	If either registry key does not exist and this parameter is not specified, the report will
	not contain a Company Name on the cover page.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	Only Word 2010, 2013 and 2016 are supported.
	(default cover pages in Word en-US)
	
	Valid input is:
		Alphabet (Word 2010. Works)
		Annual (Word 2010. Doesn't work well for this report)
		Austere (Word 2010. Works)
		Austin (Word 2010/2013/2016. Doesn't work in 2013 or 2016, mostly works in 2010 but 
						Subtitle/Subject & Author fields need to be moved 
						after title box is moved up)
		Banded (Word 2013/2016. Works)
		Conservative (Word 2010. Works)
		Contrast (Word 2010. Works)
		Cubicles (Word 2010. Works)
		Exposure (Word 2010. Works if you like looking sideways)
		Facet (Word 2013/2016. Works)
		Filigree (Word 2013/2016. Works)
		Grid (Word 2010/2013/2016. Works in 2010)
		Integral (Word 2013/2016. Works)
		Ion (Dark) (Word 2013/2016. Top date doesn't fit; box needs to be manually resized or font 
						changed to 8 point)
		Ion (Light) (Word 2013/2016. Top date doesn't fit; box needs to be manually resized or font 
						changed to 8 point)
		Mod (Word 2010. Works)
		Motion (Word 2010/2013/2016. Works if top date is manually changed to 36 point)
		Newsprint (Word 2010. Works but date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2010. Works)
		Puzzle (Word 2010. Top date doesn't fit; box needs to be manually resized or font 
					changed to 14 point)
		Retrospect (Word 2013/2016. Works)
		Semaphore (Word 2013/2016. Works)
		Sideline (Word 2010/2013/2016. Doesn't work in 2013 or 2016, works in 2010)
		Slice (Dark) (Word 2013/2016. Doesn't work)
		Slice (Light) (Word 2013/2016. Doesn't work)
		Stacks (Word 2010. Works)
		Tiles (Word 2010. Date doesn't fit unless changed to 26 point)
		Transcend (Word 2010. Works)
		ViewMaster (Word 2013/2016. Works)
		Whisp (Word 2013/2016. Works)
		
	Default value is Sideline.
	This parameter has an alias of CP.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER UserName
	User name to use for the Cover Page and Footer.
	Default value is contained in $env:username
	This parameter has an alias of UN.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER Export
    Runs this script gathering all required data from PowerCLI as normal, then exporting data to XML files in the .\Export directory
    Once the export is completed, it can be copied offline to be run later with the -Import paramater
    This parameter overrides all other output formats
.PARAMETER Import
    Runs this script gathering all required data from a previously run Export
    Export directory must be present in the same directory as the script itself
    Does not require PowerCLI or a VIServerName to run in Import mode
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	This parameter is disabled by default.
	The PDF file is roughly 5X to 10X larger than the DOCX file.
	This parameter requires Microsoft Word to be installed.
	This parameter uses the Word SaveAs PDF capability.
.PARAMETER Text
	Creates a formatted text file with a .txt extension.
	This parameter is disabled by default.
.PARAMETER MSWord
	SaveAs DOCX file
	This parameter is set True if no other output format is selected.
.PARAMETER HTML
	Creates an HTML file with an .html extension.
	This parameter is disabled by default.
.PARAMETER Full
	Runs a full inventory for the Hosts, clusters, resoure pools, networking and virtual machines.
	This parameter is disabled by default - only a summary is run when this parameter is not specified.
.PARAMETER PCLICustom
    Prompts user to locate the PowerCLI Scripts directory in a non-default installation
    This parameter is disabled by default
.PARAMETER Chart
    This parameter is still beta and is disabled by default
    Gathers data from VMware stats to build performance graphs for hosts and VMs
    DOTNET chart controls are required
.PARAMETER Issues
    This parameter is still beta and is disabled by default
    Gathers basic summary data as well as specific issues data with the idea to be run on a set schedule
    This parameter does not currently support Import\Export
.PARAMETER AddDateTime
	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2016 at 6PM is 2016-06-01_1800.
	Output filename will be ReportName_2016-06-01_1800.docx (or .pdf).
	This parameter is disabled by default.
.PARAMETER Folder
	Specifies the optional output folder to save the output report. 
.PARAMETER SmtpServer
	Specifies the optional email server to send the output report. 
.PARAMETER SmtpPort
	Specifies the SMTP port. 
	Default is 25.
.PARAMETER UseSSL
	Specifies whether to use SSL for the SmtpServer.
	Default is False.
.PARAMETER From
	Specifies the username for the From email address.
	If SmtpServer is used, this is a required parameter.
.PARAMETER To
	Specifies the username for the To email address.
	If SmtpServer is used, this is a required parameter.
.PARAMETER Dev
	Clears errors at the beginning of the script.
	Outputs all errors to a text file at the end of the script.
	
	This is used when the script developer requests more troubleshooting data.
	Text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
.PARAMETER ScriptInfo
	Outputs information about the script to a text file.
	Text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
	This parameter has an alias of SI.
.EXAMPLE
	PS C:\PSScript > .\VMware_Inventory.ps1
	
	Will use all default values and prompt for vCenter Server.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Jacob Rutski" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Jacob Rutski"
	$env:username = Administrator

	Jacob Rutski for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\VMware_Inventory.ps1 -VIServerName testvc.lab.com
	
	Will use all default values and use testvc.lab.com as the vCenter Server.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Jacob Rutski" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Jacob Rutski"
	$env:username = Administrator

	Jacob Rutski for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\VMware_Inventory.ps1 -PDF -VIServerName testvc.lab.com
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Jacob Rutski" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Jacob Rutski"
	$env:username = Administrator

	Jacob Rutski for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\VMware_Inventory.ps1 -TEXT -VIServerName testvc.lab.com

	This parameter will output a basic txt file - this output is significantly limited; HTML is recommended
	
	Will use all default values and save the document as a formatted text file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Jacob Rutski" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Jacob Rutski"
	$env:username = Administrator

	Jacob Rutski for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\VMware_Inventory.ps1 -HTML -VIServerName testvc.lab.com

	This parameter will output an HTML summary output
	
	Will use all default values and save the document as an HTML file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Jacob Rutski" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Jacob Rutski"
	$env:username = Administrator

	Jacob Rutski for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\VMware_Inventory.ps1 -Full -VIServerName testvc.lab.com
	
	Creates a full inventory of the VMware environment. *Note: a full report will take a considerable amount of time to generate.
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Jacob Rutski" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Jacob Rutski"
	$env:username = Administrator

	Jacob Rutski for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\VMware_Inventory.ps1 -PDF -Full -VIServerName testvc.lab.com
	
	Creates a full inventory of the VMware environment. *Note: a full report will take a considerable amount of time to generate.
	Will use all Default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Jacob Rutski" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Jacob Rutski"
	$env:username = Administrator

	Jacob Rutski for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript .\VMware_Inventory.ps1 -CompanyName "SeriousTek" -CoverPage "Mod" -UserName "Jacob Rutski" -VIServerName testvc.lab.com

	Will use:
		Jacob Rutski Consulting for the Company Name.
		Mod for the Cover Page format.
		Jacob Rutski for the User Name.
.EXAMPLE
    PS C:\PSScript .\VMware_Inventory.ps1 -Export -VIServerName testvc.lab.com

	Will use all default values and use testvc.lab.com as the vCenter Server.
    Script will output all data to XML files in the .\Export directory created
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Jacob Rutski" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Jacob Rutski"
	$env:username = Administrator

	Jacob Rutski for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript .\VMware_Inventory.ps1 -CN "SeriousTek" -CP "Mod" -UN "Jacob Rutski" -VIServerName testvc.lab.com

	Will use:
		Jacob Rutski Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Jacob Rutski for the User Name (alias UN).
.EXAMPLE
	PS C:\PSScript > .\VMware_Inventory.ps1 -AddDateTime -VIServerName testvc.lab.com
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Jacob Rutski" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Jacob Rutski"
	$env:username = Administrator

	Jacob Rutski for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2016 at 6PM is 2016-06-01_1800.
	Output filename will be vCenterServer_2016-06-01_1800.docx
.EXAMPLE
	PS C:\PSScript > .\VMware_Inventory.ps1 -PDF -AddDateTime -VIServerName testvc.lab.com
	
	Will use all Default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Jacob Rutski" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Jacob Rutski"
	$env:username = Administrator

	Jacob Rutski for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2016 at 6PM is 2016-06-01_1800.
	Output filename will be vCenterServerSiteName_2016-06-01_1800.pdf
.EXAMPLE
	PS C:\PSScript > .\VMware_Inventory.ps1 -Folder \\FileServer\ShareName -VIServerName testvc.lab.com
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Jacob Rutski" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Jacob Rutski"
	$env:username = Administrator

	Jacob Rutski for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Output file will be saved in the path \\FileServer\ShareName
.EXAMPLE
	PS C:\PSScript > .\VMware_Inventory.ps1 -SmtpServer mail.domain.tld -From XDAdmin@domain.tld -To ITGroup@domain.tld -VIServerName testvc.lab.com
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Jacob Rutski" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Jacob Rutski"
	$env:username = Administrator

	Jacob Rutski for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Script will use the email server mail.domain.tld, sending from XDAdmin@domain.tld, sending to ITGroup@domain.tld.
	Script will use the default SMTP port 25 and will not use SSL.
	If the current user's credentials are not valid to send email, the user will be prompted to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\VMware_Inventory.ps1 -SmtpServer smtp.office365.com -SmtpPort 587 -UseSSL -From admin@serioustek.net -To ITGroup@CarlWebster.com -VIServerName testvc.lab.com
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Jacob Rutski" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Jacob Rutski"
	$env:username = Administrator

	Jacob Rutski for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Script will use the email server smtp.office365.com on port 587 using SSL, sending from admin@serioustek.net, sending to ITGroup@carlwebster.com.
	If the current user's credentials are not valid to send email, the user will be prompted to enter valid credentials.
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  
	This script creates a Word, PDF, Formatted Text or HTML document.
.NOTES
	NAME: VMware_Inventory.ps1
	VERSION: 1.2
	AUTHOR: Jacob Rutski and Carl Webster, Sr. Solutions Architect Choice Solutions
	LASTEDIT: February 13, 2017
#>

#endregion

#region script parameters
#thanks to @jeffwouters and Michael B. Smith for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "Word") ]

Param(
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Switch]$MSWord=$False,

	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Switch]$PDF=$False,

	[parameter(ParameterSetName="Text",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Switch]$Text=$False,

	[parameter(ParameterSetName="HTML",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Switch]$HTML=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$AddDateTime=$False,
	
	[parameter(Mandatory=$False)] 
	[string]$Folder="",

    [parameter(Mandatory=$False)]
    [Alias("VC")]
    [ValidateNotNullOrEmpty()]
    [string]$VIServerName="",
	
	[parameter(Mandatory=$False)] 
	[Switch]$Full=$False,	

    [parameter(Mandatory=$False)]
    [Switch]$PCLICustom=$False,

    [parameter(Mandatory=$False)]
    [Switch]$Chart=$False,

    [parameter(Mandatory=$False)]
    [Switch]$Import=$False,

    [parameter(Mandatory=$False)]
    [Switch]$Export=$False,

    [parameter(Mandatory=$False)]
    [Switch]$Issues=$False,
	
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username,

	[parameter(ParameterSetName="SMTP",Mandatory=$True)] 
	[string]$SmtpServer="",

	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[int]$SmtpPort=25,

	[parameter(ParameterSetName="SMTP",Mandatory=$False)] 
	[switch]$UseSSL=$False,

	[parameter(ParameterSetName="SMTP",Mandatory=$True)] 
	[string]$From="",

	[parameter(ParameterSetName="SMTP",Mandatory=$True)] 
	[string]$To="",
	
	[parameter(Mandatory=$False)] 
	[Switch]$Dev=$False,
	
	[parameter(Mandatory=$False)] 
	[Alias("SI")]
	[Switch]$ScriptInfo=$False
	
	)
#endregion

#region script change log	
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#Created on June 1, 2014

#HTML functions and sample text contributed by Ken Avram October 2014
#HTML Functions FormatHTMLTable and AddHTMLTable modified by Jake Rutski May 2015
#Organized functions into logical units 16-Oct-2014
#Added regions 16-Oct-2014

#VMware vCenter inventory
#Jacob Rutski
#jake@serioustek.net
#http://blogs.serioustek.net
#@JRutski on Twitter
#Created on November 3rd, 2014
#
#Version 0.2
#-Added SSH service status, syslog log directory on hosts
#-Added VMware email settings, global settings section
#-Added VM Snapshot count
#-Fix for multiple IPs on VM
#
#Version 0.3
#-Any Gets used more than once made global
#-Fixed empty cluster
#-Finished text formatted output (no summary, compressed tables)
#-Added NTP service, licensing, summary page, check for PowerCLI version
#
#Version 0.4
#-Added heatmaps for summary tables; host block storage connections; basic DVSwitching support
#-Fixed multi column table width; fixed 32\64 OS path to PCLI
#-Set summary to default, added -Full parameter for full inventory
#-Swapped table formats for host and standard vSwitches
#
#Version 1.0
#-Fixed Get-Advanced parameters
#-Added Heatmap legend table, DSN for Windows vCenter, left-aligned tables, vCenter server version
#
#Version 1.1
#-Fix for help text region tags, fixes from template script for save as PDF, fix for memory heatmap
#-Added vCenter plugins
#
#Version 1.2
#-Added Import and Export functionality to output all data to XML that can be taken offline to generate a document at a later time
#
#Version 1.3
#-Beta chart support for performance graphs
#-Support for PowerCLI 6.0
#
#Version 1.4
#-Reworked HTML general and table functions
#-Full HTML output now functional
#-Added fix for closing Word with PDF file
#
#Version 1.5
#-Added vCenter permissions and non-standard roles
#-Added DRS Rules and Groups
#
#Version 1.5.1
#-Cleaned up some extra PCLI calls - set to variables
#-Removed almost all of the extra PCLI verbose messages - Thanks @carlwebster!!
#-Set Issues parameter to disable full run
#
#Version 1.5.2 5-Oct-2015
#	Added support for Word 2016
#
#Version 1.6
#-Added several advanced settings for VMs and VMHosts
#-Updated to ScriptTemplate 21-Feb-2016
#
#Version 1.61 Apr 21, 2016
#-Fixed title and subtitle for the Word/PDF cover page
#
#Version 1.62 19-Aug-2016
#	Fixed several misspelled words
#
#Version 1.63
#	Add support for the -Dev and -ScriptInfo parameters
#	Update the ShowScriptOptions function with all script parameters
#	Add Break statements to most Switch statements
#
#Version 1.7 22-Oct-2016
#	Added support for PowerCLI installed in non-default locations
#	Fixed formatting issues with HTML output
#	Sort Guest Volume Paths by drive letter
#
#Version 1.71 9-Nov-2016
#	Added Chinese language support
#	Fixed HTMLHeatMap
#	Fixed PWD for save path issue when importing PCLI back to C:\
#	Prompt to disconnect if PCLI is already connected
#
#endregion

#region initial variable testing and setup
Set-StrictMode -Version 2

#force  on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
$SaveEAPreference = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'

If($Null -eq $PDF)
{
	$PDF = $False
}
If($Null -eq $Text)
{
	$Text = $False
}
If($Null -eq $MSWord)
{
	$MSWord = $False
}
If($Null -eq $HTML)
{
	$HTML = $False
}
If($Null -eq $AddDateTime)
{
	$AddDateTime = $False
}
If($Full -eq $Null)
{
	$Full = $False
}
If($Chart -eq $Null)
{
	$Chart = $False
}
If($Null -eq $Folder)
{
	$Folder = ""
}
If($Null -eq $SmtpServer)
{
	$SmtpServer = ""
}
If($Null -eq $SmtpPort)
{
	$SmtpPort = 25
}
If($Null -eq $UseSSL)
{
	$UseSSL = $False
}
If($Null -eq $From)
{
	$From = ""
}
If($Null -eq $To)
{
	$To = ""
}
If($Null -eq $Dev)
{
	$Dev = $False
}
If($Null -eq $ScriptInfo)
{
	$ScriptInfo = $False
}

If(!(Test-Path Variable:PDF))
{
	$PDF = $False
}
If(!(Test-Path Variable:Text))
{
	$Text = $False
}
If(!(Test-Path Variable:MSWord))
{
	$MSWord = $False
}
If(!(Test-Path Variable:HTML))
{
	$HTML = $False
}
If(!(Test-Path Variable:AddDateTime))
{
	$AddDateTime = $False
}
If(!(Test-Path Variable:Full))
{
	$Full = $False
}
If(!(Test-Path Variable:Import))
{
	$Import = $False
}
If(!(Test-Path Variable:Export))
{
	$Export = $False
}
If(!(Test-Path Variable:Folder))
{
	$Folder = ""
}
If(!(Test-Path Variable:SmtpServer))
{
	$SmtpServer = ""
}
If(!(Test-Path Variable:SmtpPort))
{
	$SmtpPort = 25
}
If(!(Test-Path Variable:UseSSL))
{
	$UseSSL = $False
}
If(!(Test-Path Variable:From))
{
	$From = ""
}
If(!(Test-Path Variable:To))
{
	$To = ""
}
If(!(Test-Path Variable:Dev))
{
	$Dev = $False
}
If(!(Test-Path Variable:ScriptInfo))
{
	$ScriptInfo = $False
}

If($Dev)
{
	$Error.Clear()
	$Script:DevErrorFile = "$($pwd.Path)\VMwareInventoryScriptErrors_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
}

If($Null -eq $MSWord)
{
	If($Text -or $HTML -or $PDF)
	{
		$MSWord = $False
	}
	Else
	{
		$MSWord = $True
	}
}

If($MSWord -eq $False -and $PDF -eq $False -and $Text -eq $False -and $HTML -eq $False)
{
	$MSWord = $True
}

Write-Verbose "$(Get-Date): Testing output parameters"

If($MSWord)
{
	Write-Verbose "$(Get-Date): MSWord is set"
}
ElseIf($PDF)
{
	Write-Verbose "$(Get-Date): PDF is set"
}
ElseIf($Text)
{
	Write-Verbose "$(Get-Date): Text is set"
}
ElseIf($HTML)
{
	Write-Verbose "$(Get-Date): HTML is set"
}
Else
{
	$ErrorActionPreference = $SaveEAPreference
	Write-Verbose "$(Get-Date): Unable to determine output parameter"
	If($Null -eq $MSWord)
	{
		Write-Verbose "$(Get-Date): MSWord is Null"
	}
	ElseIf($Null -eq $PDF)
	{
		Write-Verbose "$(Get-Date): PDF is Null"
	}
	ElseIf($Null -eq $Text)
	{
		Write-Verbose "$(Get-Date): Text is Null"
	}
	ElseIf($Null -eq $HTML)
	{
		Write-Verbose "$(Get-Date): HTML is Null"
	}
	Else
	{
		Write-Verbose "$(Get-Date): MSWord is $($MSWord)"
		Write-Verbose "$(Get-Date): PDF is $($PDF)"
		Write-Verbose "$(Get-Date): Text is $($Text)"
		Write-Verbose "$(Get-Date): HTML is $($HTML)"
	}
	Write-Error "Unable to determine output parameter.  Script cannot continue"
	Exit
}

If($Issues)
{
    Write-Verbose "$(Get-Date): Issues is set"
    $Full = $False
    $Import = $False
    $Export = $False
}
If($Full)
{
    Write-Warning ""
    Write-Warning "Full-Run is set. This will create a full VMware inventory and will take a significant amount of time."
    Write-Warning ""
}

If($Export)
{
    Write-Warning ""
    Write-Warning "Export is set - Script will output to XML for later use, overriding any other output variables."
    Write-Warning ""

    $MSWord = $False
    $PDF = $False
    $Text = $False
    $HTML = $False
}

If(!($VIServerName) -and !($Import))
{
    $VIServerName = Read-Host 'Please enter the FQDN of your vCenter server'
}

If($Folder -ne "")
{
	Write-Verbose "$(Get-Date): Testing folder path"
	#does it exist
	If(Test-Path $Folder -EA 0)
	{
		#it exists, now check to see if it is a folder and not a file
		If(Test-Path $Folder -pathType Container -EA 0)
		{
			#it exists and it is a folder
			Write-Verbose "$(Get-Date): Folder path $Folder exists and is a folder"
		}
		Else
		{
			#it exists but it is a file not a folder
			Write-Error "Folder $Folder is a file, not a folder.  Script cannot continue"
			Exit
		}
	}
	Else
	{
		#does not exist
		Write-Error "Folder $Folder does not exist.  Script cannot continue"
		Exit
	}
}

#endregion

#region initialize variables for word html and text
[string]$Script:RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption

If($MSWord -or $PDF)
{
	#try and fix the issue with the $CompanyName variable
	$Script:CoName = $CompanyName
	Write-Verbose "$(Get-Date): CoName is $($Script:CoName)"
	
	#the following values were attained from 
	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
	#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
	[int]$wdAlignPageNumberRight = 2
	[long]$wdColorGray15 = 14277081
	[long]$wdColorGray05 = 15987699 
	[int]$wdMove = 0
	[int]$wdSeekMainDocument = 0
	[int]$wdSeekPrimaryFooter = 4
	[int]$wdStory = 6
	[int]$wdColorRed = 255
	[int]$wdColorBlack = 0
	[int]$wdWord2007 = 12
	[int]$wdWord2010 = 14
	[int]$wdWord2013 = 15
	[int]$wdWord2016 = 16
	[int]$wdFormatDocumentDefault = 16
	[int]$wdFormatPDF = 17
	#http://blogs.technet.com/b/heyscriptingguy/archive/2006/03/01/how-can-i-right-align-a-single-column-in-a-word-table.aspx
	#http://msdn.microsoft.com/en-us/library/office/ff835817%28v=office.15%29.aspx
	[int]$wdAlignParagraphLeft = 0
	[int]$wdAlignParagraphCenter = 1
	[int]$wdAlignParagraphRight = 2
	#http://msdn.microsoft.com/en-us/library/office/ff193345%28v=office.15%29.aspx
	[int]$wdCellAlignVerticalTop = 0
	[int]$wdCellAlignVerticalCenter = 1
	[int]$wdCellAlignVerticalBottom = 2
	#http://msdn.microsoft.com/en-us/library/office/ff844856%28v=office.15%29.aspx
	[int]$wdAutoFitFixed = 0
	[int]$wdAutoFitContent = 1
	[int]$wdAutoFitWindow = 2
	#http://msdn.microsoft.com/en-us/library/office/ff821928%28v=office.15%29.aspx
	[int]$wdAdjustNone = 0
	[int]$wdAdjustProportional = 1
	[int]$wdAdjustFirstColumn = 2
	[int]$wdAdjustSameWidth = 3

	[int]$PointsPerTabStop = 36
	[int]$Indent0TabStops = 0 * $PointsPerTabStop
	[int]$Indent1TabStops = 1 * $PointsPerTabStop
	[int]$Indent2TabStops = 2 * $PointsPerTabStop
	[int]$Indent3TabStops = 3 * $PointsPerTabStop
	[int]$Indent4TabStops = 4 * $PointsPerTabStop

	# http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
	[int]$wdStyleHeading1 = -2
	[int]$wdStyleHeading2 = -3
	[int]$wdStyleHeading3 = -4
	[int]$wdStyleHeading4 = -5
	[int]$wdStyleNoSpacing = -158
	[int]$wdTableGrid = -155

	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/org/codehaus/groovy/scriptom/tlb/office/word/WdLineStyle.html
	[int]$wdLineStyleNone = 0
	[int]$wdLineStyleSingle = 1

	[int]$wdHeadingFormatTrue = -1
	[int]$wdHeadingFormatFalse = 0 
}

If($HTML)
{
    Set htmlredmask         -Option AllScope -Value "#FF0000" 4>$Null
    Set htmlcyanmask        -Option AllScope -Value "#00FFFF" 4>$Null
    Set htmlbluemask        -Option AllScope -Value "#0000FF" 4>$Null
    Set htmldarkbluemask    -Option AllScope -Value "#0000A0" 4>$Null
    Set htmllightbluemask   -Option AllScope -Value "#ADD8E6" 4>$Null
    Set htmlpurplemask      -Option AllScope -Value "#800080" 4>$Null
    Set htmlyellowmask      -Option AllScope -Value "#FFFF00" 4>$Null
    Set htmllimemask        -Option AllScope -Value "#00FF00" 4>$Null
    Set htmlmagentamask     -Option AllScope -Value "#FF00FF" 4>$Null
    Set htmlwhitemask       -Option AllScope -Value "#FFFFFF" 4>$Null
    Set htmlsilvermask      -Option AllScope -Value "#C0C0C0" 4>$Null
    Set htmlgraymask        -Option AllScope -Value "#808080" 4>$Null
    Set htmlblackmask       -Option AllScope -Value "#000000" 4>$Null
    Set htmlorangemask      -Option AllScope -Value "#FFA500" 4>$Null
    Set htmlmaroonmask      -Option AllScope -Value "#800000" 4>$Null
    Set htmlgreenmask       -Option AllScope -Value "#008000" 4>$Null
    Set htmlolivemask       -Option AllScope -Value "#808000" 4>$Null

    Set htmlbold        -Option AllScope -Value 1 4>$Null
    Set htmlitalics     -Option AllScope -Value 2 4>$Null
    Set htmlred         -Option AllScope -Value 4 4>$Null
    Set htmlcyan        -Option AllScope -Value 8 4>$Null
    Set htmlblue        -Option AllScope -Value 16 4>$Null
    Set htmldarkblue    -Option AllScope -Value 32 4>$Null
    Set htmllightblue   -Option AllScope -Value 64 4>$Null
    Set htmlpurple      -Option AllScope -Value 128 4>$Null
    Set htmlyellow      -Option AllScope -Value 256 4>$Null
    Set htmllime        -Option AllScope -Value 512 4>$Null
    Set htmlmagenta     -Option AllScope -Value 1024 4>$Null
    Set htmlwhite       -Option AllScope -Value 2048 4>$Null
    Set htmlsilver      -Option AllScope -Value 4096 4>$Null
    Set htmlgray        -Option AllScope -Value 8192 4>$Null
    Set htmlolive       -Option AllScope -Value 16384 4>$Null
    Set htmlorange      -Option AllScope -Value 32768 4>$Null
    Set htmlmaroon      -Option AllScope -Value 65536 4>$Null
    Set htmlgreen       -Option AllScope -Value 131072 4>$Null
    Set htmlblack       -Option AllScope -Value 262144 4>$Null
}

If($TEXT)
{
	$global:output = ""
}
#endregion

#region word specific functions
Function SetWordHashTable
{
	Param([string]$CultureCode)

	#optimized by Michael B. SMith
	
	# DE and FR translations for Word 2010 by Vladimir Radojevic
	# Vladimir.Radojevic@Commerzreal.com

	# DA translations for Word 2010 by Thomas Daugaard
	# Citrix Infrastructure Specialist at edgemo A/S

	# CA translations by Javier Sanchez 
	# CEO & Founder 101 Consulting

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese
	
	[string]$toc = $(
		Switch ($CultureCode)
		{
			'ca-'	{ 'Taula automática 2'; Break }
			'da-'	{ 'Automatisk tabel 2'; Break }
			'de-'	{ 'Automatische Tabelle 2'; Break }
			'en-'	{ 'Automatic Table 2'; Break }
			'es-'	{ 'Tabla automática 2'; Break }
			'fi-'	{ 'Automaattinen taulukko 2'; Break }
			'fr-'	{ 'Table automatique 2'; Break } #changed 13-feb-2017 david roquier and samuel legrand
			'nb-'	{ 'Automatisk tabell 2'; Break }
			'nl-'	{ 'Automatische inhoudsopgave 2'; Break }
			'pt-'	{ 'Sumário Automático 2'; Break }
			'sv-'	{ 'Automatisk innehållsförteckning2'; Break }
			'zh-'	{ '自动目录 2'; Break }
		}
	)

	$Script:myHash                      = @{}
	$Script:myHash.Word_TableOfContents = $toc
	$Script:myHash.Word_NoSpacing       = $wdStyleNoSpacing
	$Script:myHash.Word_Heading1        = $wdStyleheading1
	$Script:myHash.Word_Heading2        = $wdStyleheading2
	$Script:myHash.Word_Heading3        = $wdStyleheading3
	$Script:myHash.Word_Heading4        = $wdStyleheading4
	$Script:myHash.Word_TableGrid       = $wdTableGrid
}

Function GetCulture
{
	Param([int]$WordValue)
	
	#codes obtained from http://support.microsoft.com/kb/221435
	#http://msdn.microsoft.com/en-us/library/bb213877(v=office.12).aspx
	$CatalanArray = 1027
	$ChineseArray = 2052,3076,5124,4100
	$DanishArray = 1030
	$DutchArray = 2067, 1043
	$EnglishArray = 3081, 10249, 4105, 9225, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 1033, 12297
	$FinnishArray = 1035
	$FrenchArray = 2060, 1036, 11276, 3084, 12300, 5132, 13324, 6156, 8204, 10252, 7180, 9228, 4108
	$GermanArray = 1031, 3079, 5127, 4103, 2055
	$NorwegianArray = 1044, 2068
	$PortugueseArray = 1046, 2070
	$SpanishArray = 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 19466, 6154, 15370, 10250, 20490, 3082, 14346, 8202
	$SwedishArray = 1053, 2077

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese

	Switch ($WordValue)
	{
		{$CatalanArray -contains $_} {$CultureCode = "ca-"}
		{$ChineseArray -contains $_} {$CultureCode = "zh-"}
		{$DanishArray -contains $_} {$CultureCode = "da-"}
		{$DutchArray -contains $_} {$CultureCode = "nl-"}
		{$EnglishArray -contains $_} {$CultureCode = "en-"}
		{$FinnishArray -contains $_} {$CultureCode = "fi-"}
		{$FrenchArray -contains $_} {$CultureCode = "fr-"}
		{$GermanArray -contains $_} {$CultureCode = "de-"}
		{$NorwegianArray -contains $_} {$CultureCode = "nb-"}
		{$PortugueseArray -contains $_} {$CultureCode = "pt-"}
		{$SpanishArray -contains $_} {$CultureCode = "es-"}
		{$SwedishArray -contains $_} {$CultureCode = "sv-"}
		Default {$CultureCode = "en-"}
	}
	
	Return $CultureCode
}

Function ValidateCoverPage
{
	Param([int]$xWordVersion, [string]$xCP, [string]$CultureCode)
	
	$xArray = ""
	
	Switch ($CultureCode)
	{
		'ca-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Anual", "Austin", "Conservador",
					"Contrast", "Cubicles", "Diplomàtic", "Exposició",
					"Línia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
					"Perspectiva", "Piles", "Quadrícula", "Sobri",
					"Transcendir", "Trencaclosques")
				}
			}

		'da-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevægElse", "Brusen", "Facet", "Filigran", 
					"Gitter", "Integral", "Ion (lys)", "Ion (mørk)", 
					"Retro", "Semafor", "Sidelinje", "Stribet", 
					"Udsnit (lys)", "Udsnit (mørk)", "Visningsmaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("BevægElse", "Brusen", "Ion (lys)", "Filigran",
					"Retro", "Semafor", "Visningsmaster", "Integral",
					"Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
					"Udsnit (mørk)", "Ion (mørk)", "Austin")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("BevægElse", "Moderat", "Perspektiv", "Firkanter",
					"Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "Gåde",
					"Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
					"Nålestribet", "Årlig", "Avispapir", "Tradionel")
				}
			}

		'de-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Bewegung", "Facette", "Filigran", 
					"Gebändert", "Integral", "Ion (dunkel)", "Ion (hell)", 
					"Pfiff", "Randlinie", "Raster", "Rückblick", 
					"Segment (dunkel)", "Segment (hell)", "Semaphor", 
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
					"Raster", "Ion (dunkel)", "Filigran", "Rückblick", "Pfiff",
					"ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
					"Randlinie", "Austin", "Integral", "Facette")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
					"Herausgestellt", "Jährlich", "Kacheln", "Kontrast", "Kubistisch",
					"Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie",
					"Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
				}
			}

		'en-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
					"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
					"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
					"Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
					"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
					"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
				}
			}

		'es-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Con bandas", "Cortar (oscuro)", "Cuadrícula", 
					"Whisp", "Faceta", "Filigrana", "Integral", "Ion (claro)", 
					"Ion (oscuro)", "Línea lateral", "Movimiento", "Retrospectiva", 
					"Semáforo", "Slice (luz)", "Vista principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
					"Slice (luz)", "Faceta", "Semáforo", "Retrospectiva", "Cuadrícula",
					"Movimiento", "Cortar (oscuro)", "Línea lateral", "Ion (oscuro)",
					"Ion (claro)", "Integral", "Con bandas")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
					"Contraste", "Cuadrícula", "Cubículos", "Exposición", "Línea lateral",
					"Moderno", "Mosaicos", "Movimiento", "Papel periódico",
					"Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
				}
			}

		'fi-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kuiskaus", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kiehkura", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti",
					"Laatikot", "Liike", "Liituraita", "Mod", "Osittain peitossa",
					"Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko",
					"Ruudut", "Sanomalehtipaperi", "Sivussa", "Vuotuinen", "Ylitys")
				}
			}

		'fr-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("À bandes", "Austin", "Facette", "Filigrane", 
					"Guide", "Intégrale", "Ion (clair)", "Ion (foncé)", 
					"Lignes latérales", "Quadrillage", "Rétrospective", "Secteur (clair)", 
					"Secteur (foncé)", "Sémaphore", "ViewMaster", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annuel", "Austère", "Austin", 
					"Blocs empilés", "Classique", "Contraste", "Emplacements de bureau", 
					"Exposition", "Guide", "Ligne latérale", "Moderne", 
					"Mosaïques", "Mots croisés", "Papier journal", "Perspective",
					"Quadrillage", "Rayures fines", "Transcendant")
				}
			}

		'nb-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran",
					"Integral", "Ion (lys)", "Ion (mørk)", "Retrospekt", "Rutenett",
					"Sektor (lys)", "Sektor (mørk)", "Semafor", "Sidelinje", "Stripet",
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Årlig", "Avistrykk", "Austin", "Avlukker",
					"BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
					"Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje",
					"Smale striper", "Stabler", "Transcenderende")
				}
			}

		'nl-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept",
					"Integraal", "Ion (donker)", "Ion (licht)", "Raster",
					"Segment (Light)", "Semafoor", "Slice (donker)", "Spriet",
					"Terugblik", "Terzijde", "ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden",
					"Beweging", "Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks",
					"Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief",
					"Puzzel", "Raster", "Stapels",
					"Tegels", "Terzijde")
				}
			}

		'pt-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Animação", "Austin", "Em Tiras", "Exibição Mestra",
					"Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana", 
					"Grade", "Integral", "Íon (Claro)", "Íon (Escuro)", "Linha Lateral",
					"Retrospectiva", "Semáforo")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Austin", "Baias",
					"Conservador", "Contraste", "Exposição", "Grade", "Ladrilhos",
					"Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
					"Quebra-cabeça", "Transcend")
				}
			}

		'sv-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
					"Jon (mörkt)", "Knippe", "Rutnät", "RörElse", "Sektor (ljus)", "Sektor (mörk)",
					"Semafor", "Sidlinje", "VisaHuvudsida", "Återblick")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabetmönster", "Austin", "Enkelt", "Exponering", "Konservativt",
					"Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "Rutnät",
					"RörElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Årligt",
					"Övergående")
				}
			}

		'zh-'	{
				If($xWordVersion -eq $wdWord2010 -or $xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ('奥斯汀', '边线型', '花丝', '怀旧', '积分',
					'离子(浅色)', '离子(深色)', '母版型', '平面', '切片(浅色)',
					'切片(深色)', '丝状', '网格', '镶边', '信号灯',
					'运动型')
				}
			}

		Default	{
					If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
					{
						$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
						"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
						"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
						"Whisp")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
						"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
						"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
					}
				}
	}
	
	If($xArray -contains $xCP)
	{
		$xArray = $Null
		Return $True
	}
	Else
	{
		$xArray = $Null
		Return $False
	}
}

Function CheckWordPrereq
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`t`tThis script directly outputs to Microsoft Word, please install Microsoft Word`n`n"
		Exit
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	[bool]$wordrunning = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}) -ne $Null
	If($wordrunning)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`tPlease close all instances of Microsoft Word before running this report.`n`n"
		Exit
	}
}

Function ValidateCompanyName
{
	[bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	If($xResult)
	{
		Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	}
	Else
	{
		$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		If($xResult)
		{
			Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		}
		Else
		{
			Return ""
		}
	}
}

Function _SetDocumentProperty 
{
	#jeff hicks
	Param([object]$Properties,[string]$Name,[string]$Value)
	#get the property object
	$prop = $properties | ForEach { 
		$propname = ""
		$propname=$_.GetType().InvokeMember("Name","GetProperty",$Null,$_,$Null)
		If($propname -eq $Name) 
		{
			Return $_
		}
	} #ForEach

	#set the value
	$Prop.GetType().InvokeMember("Value","SetProperty",$Null,$prop,$Value)
}

Function FindWordDocumentEnd
{
	#return focus to main document    
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Script:Selection.EndKey($wdStory,$wdMove) | Out-Null
}

Function SetupWord
{
	Write-Verbose "$(Get-Date): Setting up Word"
    
	# Setup word for output
	Write-Verbose "$(Get-Date): Create Word comObject."
	$Script:Word = New-Object -comobject "Word.Application" -EA 0 4>$Null
	
	If(!$? -or $Null -eq $Script:Word)
	{
		Write-Warning "The Word object could not be created.  You may need to repair your Word installation."
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tThe Word object could not be created.  You may need to repair your Word installation.`n`n`t`tScript cannot continue.`n`n"
		Exit
	}

	Write-Verbose "$(Get-Date): Determine Word language value"
	If( ( validStateProp $Script:Word Language Value__ ) )
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language.Value__
	}
	Else
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language
	}

	If(!($Script:WordLanguageValue -gt -1))
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tUnable to determine the Word language value.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}
	Write-Verbose "$(Get-Date): Word language value is $($Script:WordLanguageValue)"
	
	$Script:WordCultureCode = GetCulture $Script:WordLanguageValue
	
	SetWordHashTable $Script:WordCultureCode
	
	[int]$Script:WordVersion = [int]$Script:Word.Version
	If($Script:WordVersion -eq $wdWord2016)
	{
		$Script:WordProduct = "Word 2016"
	}
	ElseIf($Script:WordVersion -eq $wdWord2013)
	{
		$Script:WordProduct = "Word 2013"
	}
	ElseIf($Script:WordVersion -eq $wdWord2010)
	{
		$Script:WordProduct = "Word 2010"
	}
	ElseIf($Script:WordVersion -eq $wdWord2007)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tMicrosoft Word 2007 is no longer supported.`n`n`t`tScript will end.`n`n"
		AbortScript
	}
	Else
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tYou are running an untested or unsupported version of Microsoft Word.`n`n`t`tScript will end.`n`n`t`tPlease send info on your version of Word to webster@carlwebster.com`n`n"
		AbortScript
	}

	#only validate CompanyName if the field is blank
	If([String]::IsNullOrEmpty($Script:CoName))
	{
		Write-Verbose "$(Get-Date): Company name is blank.  Retrieve company name from registry."
		$TmpName = ValidateCompanyName
		
		If([String]::IsNullOrEmpty($TmpName))
		{
			Write-Warning "`n`n`t`tCompany Name is blank so Cover Page will not show a Company Name."
			Write-Warning "`n`t`tCheck HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value."
			Write-Warning "`n`t`tYou may want to use the -CompanyName parameter if you need a Company Name on the cover page.`n`n"
		}
		Else
		{
			$Script:CoName = $TmpName
			Write-Verbose "$(Get-Date): Updated company name to $($Script:CoName)"
		}
	}

	If($Script:WordCultureCode -ne "en-")
	{
		Write-Verbose "$(Get-Date): Check Default Cover Page for $($WordCultureCode)"
		[bool]$CPChanged = $False
		Switch ($Script:WordCultureCode)
		{
			'ca-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línia lateral"
						$CPChanged = $True
					}
				}

			'da-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'de-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Randlinie"
						$CPChanged = $True
					}
				}

			'es-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línea lateral"
						$CPChanged = $True
					}
				}

			'fi-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sivussa"
						$CPChanged = $True
					}
				}

			'fr-'	{
					If($CoverPage -eq "Sideline")
					{
						If($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
						{
							$CoverPage = "Lignes latérales"
							$CPChanged = $True
						}
						Else
						{
							$CoverPage = "Ligne latérale"
							$CPChanged = $True
						}
					}
				}

			'nb-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'nl-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Terzijde"
						$CPChanged = $True
					}
				}

			'pt-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Linha Lateral"
						$CPChanged = $True
					}
				}

			'sv-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidlinje"
						$CPChanged = $True
					}
				}

			'zh-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "边线型"
						$CPChanged = $True
					}
				}
		}

		If($CPChanged)
		{
			Write-Verbose "$(Get-Date): Changed Default Cover Page from Sideline to $($CoverPage)"
		}
	}

	Write-Verbose "$(Get-Date): Validate cover page $($CoverPage) for culture code $($Script:WordCultureCode)"
	[bool]$ValidCP = $False
	
	$ValidCP = ValidateCoverPage $Script:WordVersion $CoverPage $Script:WordCultureCode
	
	If(!$ValidCP)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Verbose "$(Get-Date): Word language value $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date): Culture code $($Script:WordCultureCode)"
		Write-Error "`n`n`t`tFor $($Script:WordProduct), $($CoverPage) is not a valid Cover Page option.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	ShowScriptOptions

	$Script:Word.Visible = $False

	#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
	#using Jeff's Demo-WordReport.ps1 file for examples
	Write-Verbose "$(Get-Date): Load Word Templates"

	[bool]$Script:CoverPagesExist = $False
	[bool]$BuildingBlocksExist = $False

	$Script:Word.Templates.LoadBuildingBlocks()
	#word 2010/2013/2016
	$BuildingBlocksCollection = $Script:Word.Templates | Where {$_.name -eq "Built-In Building Blocks.dotx"}

	Write-Verbose "$(Get-Date): Attempt to load cover page $($CoverPage)"
	$part = $Null

	$BuildingBlocksCollection | 
	ForEach{
		If ($_.BuildingBlockEntries.Item($CoverPage).Name -eq $CoverPage) 
		{
			$BuildingBlocks = $_
		}
	}        

	If($Null -ne $BuildingBlocks)
	{
		$BuildingBlocksExist = $True

		Try 
		{
			$part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
		}

		Catch
		{
			$part = $Null
		}

		If($Null -ne $part)
		{
			$Script:CoverPagesExist = $True
		}
	}

	If(!$Script:CoverPagesExist)
	{
		Write-Verbose "$(Get-Date): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Warning "Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Warning "This report will not have a Cover Page."
	}

	Write-Verbose "$(Get-Date): Create empty word doc"
	$Script:Doc = $Script:Word.Documents.Add()
	If($Null -eq $Script:Doc)
	{
		Write-Verbose "$(Get-Date): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tAn empty Word document could not be created.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	$Script:Selection = $Script:Word.Selection
	If($Null -eq $Script:Selection)
	{
		Write-Verbose "$(Get-Date): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tAn unknown error happened selecting the entire Word document for default formatting options.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	#set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
	#36 = .50"
	$Script:Word.ActiveDocument.DefaultTabStop = 36

	#Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
	Write-Verbose "$(Get-Date): Disable grammar and spell checking"
	#bug reported 1-Apr-2014 by Tim Mangan
	#save current options first before turning them off
	$Script:CurrentGrammarOption = $Script:Word.Options.CheckGrammarAsYouType
	$Script:CurrentSpellingOption = $Script:Word.Options.CheckSpellingAsYouType
	$Script:Word.Options.CheckGrammarAsYouType = $False
	$Script:Word.Options.CheckSpellingAsYouType = $False

	If($BuildingBlocksExist)
	{
		#insert new page, getting ready for table of contents
		Write-Verbose "$(Get-Date): Insert new page, getting ready for table of contents"
		$part.Insert($Script:Selection.Range,$True) | Out-Null
		$Script:Selection.InsertNewPage()

		#table of contents
		Write-Verbose "$(Get-Date): Table of Contents - $($Script:MyHash.Word_TableOfContents)"
		$toc = $BuildingBlocks.BuildingBlockEntries.Item($Script:MyHash.Word_TableOfContents)
		If($Null -eq $toc)
		{
			Write-Verbose "$(Get-Date): "
			Write-Verbose "$(Get-Date): Table of Content - $($Script:MyHash.Word_TableOfContents) could not be retrieved."
			Write-Warning "This report will not have a Table of Contents."
		}
		Else
		{
			$toc.insert($Script:Selection.Range,$True) | Out-Null
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): Table of Contents are not installed."
		Write-Warning "Table of Contents are not installed so this report will not have a Table of Contents."
	}

	#set the footer
	Write-Verbose "$(Get-Date): Set the footer"
	[string]$footertext = "Report created by $username"

	#get the footer
	Write-Verbose "$(Get-Date): Get the footer and format font"
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekPrimaryFooter
	#get the footer and format font
	$footers = $Script:Doc.Sections.Last.Footers
	ForEach ($footer in $footers) 
	{
		If($footer.exists) 
		{
			$footer.range.Font.name = "Calibri"
			$footer.range.Font.size = 8
			$footer.range.Font.Italic = $True
			$footer.range.Font.Bold = $True
		}
	} #end ForEach
	Write-Verbose "$(Get-Date): Footer text"
	$Script:Selection.HeaderFooter.Range.Text = $footerText

	#add page numbering
	Write-Verbose "$(Get-Date): Add page numbering"
	$Script:Selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

	FindWordDocumentEnd
	Write-Verbose "$(Get-Date):"
	#end of Jeff Hicks 
}

Function UpdateDocumentProperties
{
	Param([string]$AbstractTitle, [string]$SubjectTitle)
	#Update document properties
	If($MSWORD -or $PDF)
	{
		If($Script:CoverPagesExist)
		{
			Write-Verbose "$(Get-Date): Set Cover Page Properties"
			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Company" $Script:CoName
			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Title" $Script:title
			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Author" $username

			_SetDocumentProperty $Script:Doc.BuiltInDocumentProperties "Subject" $SubjectTitle

			#Get the Coverpage XML part
			$cp = $Script:Doc.CustomXMLParts | Where {$_.NamespaceURI -match "coverPageProps$"}

			#get the abstract XML part
			$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "Abstract"}

			#set the text
			If([String]::IsNullOrEmpty($Script:CoName))
			{
				[string]$abstract = $AbstractTitle
			}
			Else
			{
				[string]$abstract = "$($AbstractTitle) for $($Script:CoName)"
			}

			$ab.Text = $abstract

			$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "PublishDate"}
			#set the text
			[string]$abstract = (Get-Date -Format d).ToString()
			$ab.Text = $abstract

			Write-Verbose "$(Get-Date): Update the Table of Contents"
			#update the Table of Contents
			$Script:Doc.TablesOfContents.item(1).Update()
			$cp = $Null
			$ab = $Null
			$abstract = $Null
		}
	}
}
#endregion

#region registry functions
#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This Function just gets $True or $False
Function Test-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	$key -and $Null -ne $key.GetValue($name, $Null)
}

# Gets the specified registry value or $Null if it is missing
Function Get-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	If($key)
	{
		$key.GetValue($name, $Null)
	}
	Else
	{
		$Null
	}
}
#endregion

#region word, text and html line output functions
Function line
#function created by Michael B. Smith, Exchange MVP
#@essentialexchange on Twitter
#http://TheEssentialExchange.com
#for creating the formatted text report
#created March 2011
#updated March 2014
{
	Param( [int]$tabs = 0, [string]$name = '', [string]$value = '', [string]$newline = "`r`n", [switch]$nonewline )
	While( $tabs -gt 0 ) { $Global:Output += "`t"; $tabs--; }
	If( $nonewline )
	{
		$Global:Output += $name + $value
	}
	Else
	{
		$Global:Output += $name + $value + $newline
	}
}
	
Function WriteWordLine
#Function created by Ryan Revord
#@rsrevord on Twitter
#Function created to make output to Word easy in this script
#updated 27-Mar-2014 to include font name, font size, italics and bold options
{
	Param([int]$style=0, 
	[int]$tabs = 0, 
	[string]$name = '', 
	[string]$value = '', 
	[string]$fontName=$Null,
	[int]$fontSize=0,
	[bool]$italics=$False,
	[bool]$boldface=$False,
	[Switch]$nonewline)
	
	#Build output style
	[string]$output = ""
	Switch ($style)
	{
		0 {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
		1 {$Script:Selection.Style = $Script:MyHash.Word_Heading1; Break}
		2 {$Script:Selection.Style = $Script:MyHash.Word_Heading2; Break}
		3 {$Script:Selection.Style = $Script:MyHash.Word_Heading3; Break}
		4 {$Script:Selection.Style = $Script:MyHash.Word_Heading4; Break}
		Default {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
	}
	
	#build # of tabs
	While($tabs -gt 0)
	{ 
		$output += "`t"; $tabs--; 
	}
 
	If(![String]::IsNullOrEmpty($fontName)) 
	{
		$Script:Selection.Font.name = $fontName
	} 

	If($fontSize -ne 0) 
	{
		$Script:Selection.Font.size = $fontSize
	} 
 
	If($italics -eq $True) 
	{
		$Script:Selection.Font.Italic = $True
	} 
 
	If($boldface -eq $True) 
	{
		$Script:Selection.Font.Bold = $True
	} 

	#output the rest of the parameters.
	$output += $name + $value
	$Script:Selection.TypeText($output)
 
	#test for new WriteWordLine 0.
	If($nonewline)
	{
		# Do nothing.
	} 
	Else 
	{
		$Script:Selection.TypeParagraph()
	}
}

#***********************************************************************************************************
# WriteHTMLLine
#***********************************************************************************************************

<#
.Synopsis
	Writes a line of output for HTML output
.DESCRIPTION
	This function formats an HTML line
.USAGE
	WriteHTMLLine <Style> <Tabs> <Name> <Value> <Font Name> <Font Size> <Options>

	0 for Font Size denotes using the default font size of 2 or 10 point

.EXAMPLE
	WriteHTMLLine 0 0 " "

	Writes a blank line with no style or tab stops, obviously none needed.

.EXAMPLE
	WriteHTMLLine 0 1 "This is a regular line of text indented 1 tab stops"

	Writes a line with 1 tab stop.

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in italics" "" $null 0 $htmlitalics

	Writes a line omitting font and font size and setting the italics attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold" "" $null 0 $htmlbold

	Writes a line omitting font and font size and setting the bold attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold italics" "" $null 0 ($htmlbold -bor $htmlitalics)

	Writes a line omitting font and font size and setting both italics and bold options

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in 10 point" "" $null 2  # 10 point font

	Writes a line using 10 point font

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in Courier New font" "" "Courier New" 0 

	Writes a line using Courier New Font and 0 font point size (default = 2 if set to 0)

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of RED text indented 0 tab stops with the computer name as data in 10 point Courier New bold italics: " $env:computername "Courier New" 2 ($htmlbold -bor $htmlred -bor $htmlitalics)

	Writes a line using Courier New Font with first and second string values to be used, also uses 10 point font with bold, italics and red color options set.

.NOTES

	Font Size - Unlike word, there is a limited set of font sizes that can be used in HTML.  They are:
		0 - default which actually gives it a 2 or 10 point.
		1 - 7.5 point font size
		2 - 10 point
		3 - 13.5 point
		4 - 15 point
		5 - 18 point
		6 - 24 point
		7 - 36 point
	Any number larger than 7 defaults to 7

	Style - Refers to the headers that are used with output and resemble the headers in word, 
	HTML supports headers h1-h6 and h1-h4 are more commonly used.  Unlike word, H1 will not 
	give you a blue colored font, you will have to set that yourself.

	Colors and Bold/Italics Flags are:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack       
#>

Function WriteHTMLLine
#Function created by Ken Avram
#Function created to make output to HTML easy in this script
#headings fixed 12-Oct-2016 by Webster
{
	Param([int]$style=0, 
	[int]$tabs = 0, 
	[string]$name = '', 
	[string]$value = '', 
	[string]$fontName="Calibri",
	[int]$fontSize=1,
	[int]$options=$htmlblack)


	#Build output style
	[string]$output = ""

	If([String]::IsNullOrEmpty($Name))	
	{
		$HTMLBody = "<p></p>"
	}
	Else
	{
		$color = CheckHTMLColor $options

		#build # of tabs

		While($tabs -gt 0)
		{ 
			$output += "&nbsp;&nbsp;&nbsp;&nbsp;"; $tabs--; 
		}

		$HTMLFontName = $fontName		

		$HTMLBody = ""

		If($options -band $htmlitalics) 
		{
			$HTMLBody += "<i>"
		} 

		If($options -band $htmlbold) 
		{
			$HTMLBody += "<b>"
		} 

		#output the rest of the parameters.
		$output += $name + $value

		#added by webster 12-oct-2016
		#if a heading, don't add the <br>
		If($HTMLStyle -eq "")
		{
			$HTMLBody += "<br><font face='" + $HTMLFontName + "' " + "color='" + $color + "' size='"  + $fontsize + "'>"
		}
		Else
		{
			$HTMLBody += "<font face='" + $HTMLFontName + "' " + "color='" + $color + "' size='"  + $fontsize + "'>"
		}
		
		Switch ($style)
		{
			1 {$HTMLStyle = "<h1>"; Break}
			2 {$HTMLStyle = "<h2>"; Break}
			3 {$HTMLStyle = "<h3>"; Break}
			4 {$HTMLStyle = "<h4>"; Break}
			Default {$HTMLStyle = ""; Break}
		}

		$HTMLBody += $HTMLStyle + $output

		Switch ($style)
		{
			1 {$HTMLStyle = "</h1>"; Break}
			2 {$HTMLStyle = "</h2>"; Break}
			3 {$HTMLStyle = "</h3>"; Break}
			4 {$HTMLStyle = "</h4>"; Break}
			Default {$HTMLStyle = ""; Break}
		}

		$HTMLBody += $HTMLStyle +  "</font>"

		If($options -band $htmlitalics) 
		{
			$HTMLBody += "</i>"
		} 

		If($options -band $htmlbold) 
		{
			$HTMLBody += "</b>"
		} 
	}
	
	#added by webster 12-oct-2016
	#if a heading, don't add the <br />
	If($HTMLStyle -eq "")
	{
		$HTMLBody += "<br />"
	}

	out-file -FilePath $Script:FileName1 -Append -InputObject $HTMLBody 4>$Null
}
#endregion

#region HTML table functions
#***********************************************************************************************************
# AddHTMLTable - Called from FormatHTMLTable function
# Created by Ken Avram
# modified by Jake Rutski
#***********************************************************************************************************
Function AddHTMLTable
{
	Param([string]$fontName="Calibri",
	[int]$fontSize=2,
	[int]$colCount=0,
	[int]$rowCount=0,
	[object[]]$rowInfo=@(),
	[object[]]$fixedInfo=@())

	For($rowidx = $RowIndex;$rowidx -le $rowCount;$rowidx++)
	{
		$rd = @($rowInfo[$rowidx - 2])
		$htmlbody = $htmlbody + "<tr>"
		For($columnIndex = 0; $columnIndex -lt $colCount; $columnindex+=2)
		{
			$fontitalics = $False
			$fontbold = $false
			$tmp = CheckHTMLColor $rd[$columnIndex+1]

			If($fixedInfo.Length -eq 0)
			{
				$htmlbody += "<td style=""background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}
			Else
			{
				$htmlbody += "<td style=""width:$($fixedInfo[$columnIndex/2]); background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}

			If($rd[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "<b>"
			}
			If($rd[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "<i>"
			}
			If($Null -ne $rd[$columnIndex])
			{
				$cell = $rd[$columnIndex].tostring()
				If($cell -eq " " -or $cell.length -eq 0)
				{
					$htmlbody += "&nbsp;&nbsp;&nbsp;"
				}
				Else
				{
					For($i=0;$i -lt $cell.length;$i++)
					{
						If($cell[$i] -eq " ")
						{
							$htmlbody += "&nbsp;"
						}
						If($cell[$i] -ne " ")
						{
							Break
						}
					}
					$htmlbody += $cell
				}
			}
			Else
			{
				$htmlbody += "&nbsp;&nbsp;&nbsp;"
			}
			If($rd[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "</b>"
			}
			If($rd[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "</i>"
			}
			$htmlbody += "</font></td>"
		}
		$htmlbody += "</tr>"
	}
	out-file -FilePath $Script:FileName1 -Append -InputObject $HTMLBody 4>$Null 
}

#***********************************************************************************************************
# FormatHTMLTable 
# Created by Ken Avram
# modified by Jake Rutski
#***********************************************************************************************************

<#
.Synopsis
	Format table for HTML output document
.DESCRIPTION
	This function formats a table for HTML from an array of strings
.PARAMETER noBorder
	If set to $true, a table will be generated without a border (border='0')
.PARAMETER noHeadCols
	This parameter should be used when generating tables without column headers
	Set this parameter equal to the number of columns in the table
.PARAMETER rowArray
	This parameter contains the row data array for the table
.PARAMETER columnArray
	This parameter contains column header data for the table
.PARAMETER fixedWidth
	This parameter contains widths for columns in pixel format ("100px") to override auto column widths
	The variable should contain a width for each column you wish to override the auto-size setting
	For example: $columnWidths = @("100px","110px","120px","130px","140px")

.USAGE
	FormatHTMLTable <Table Header> <Table Format> <Font Name> <Font Size>

.EXAMPLE
	FormatHTMLTable "Table Heading" "auto" "Calibri" 3

	This example formats a table and writes it out into an html file.  All of the parameters are optional
	defaults are used if not supplied.

	for <Table format>, the default is auto which will autofit the text into the columns and adjust to the longest text in that column.  You can also use percentage i.e. 25%
	which will take only 25% of the line and will auto word wrap the text to the next line in the column.  Also, instead of using a percentage, you can use pixels i.e. 400px.

	FormatHTMLTable "Table Heading" "auto" -rowArray $rowData -columnArray $columnData

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, column header data from $columnData and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -noHeadCols 3

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, no header, and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -fixedWidth $fixedColumns

	This example creates an HTML table with a heading of 'Table Heading, no header, row data from $rowData, and fixed columns defined by $fixedColumns

.NOTES
	In order to use the formatted table it first has to be loaded with data.  Examples below will show how to load the table:

	First, initialize the table array

	$rowdata = @()

	Then Load the array.  If you are using column headers then load those into the column headers array, otherwise the first line of the table goes into the column headers array
	and the second and subsequent lines go into the $rowdata table as shown below:

	$columnHeaders = @('Display Name',($htmlsilver -bor $htmlbold),'Status',($htmlsilver -bor $htmlbold),'Startup Type',($htmlsilver -bor $htmlbold))

	The first column is the actual name to display, the second are the attributes of the column i.e. color anded with bold or italics.  For the anding, parens are required or it will
	not format correctly.

	This is following by adding rowdata as shown below.  As more columns are added the columns will auto adjust to fit the size of the page.

	$rowdata = @()
	$columnHeaders = @("User Name",($htmlsilver -bor $htmlbold),$UserName,$htmlwhite)
	$rowdata += @(,('Save as PDF',($htmlsilver -bor $htmlbold),$PDF.ToString(),$htmlwhite))
	$rowdata += @(,('Save as TEXT',($htmlsilver -bor $htmlbold),$TEXT.ToString(),$htmlwhite))
	$rowdata += @(,('Save as WORD',($htmlsilver -bor $htmlbold),$MSWORD.ToString(),$htmlwhite))
	$rowdata += @(,('Save as HTML',($htmlsilver -bor $htmlbold),$HTML.ToString(),$htmlwhite))
	$rowdata += @(,('Add DateTime',($htmlsilver -bor $htmlbold),$AddDateTime.ToString(),$htmlwhite))
	$rowdata += @(,('Hardware Inventory',($htmlsilver -bor $htmlbold),$Hardware.ToString(),$htmlwhite))
	$rowdata += @(,('Computer Name',($htmlsilver -bor $htmlbold),$ComputerName,$htmlwhite))
	$rowdata += @(,('Filename1',($htmlsilver -bor $htmlbold),$Script:FileName1,$htmlwhite))
	$rowdata += @(,('OS Detected',($htmlsilver -bor $htmlbold),$Script:RunningOS,$htmlwhite))
	$rowdata += @(,('PSUICulture',($htmlsilver -bor $htmlbold),$PSCulture,$htmlwhite))
	$rowdata += @(,('PoSH version',($htmlsilver -bor $htmlbold),$Host.Version.ToString(),$htmlwhite))
	FormatHTMLTable "Example of Horizontal AutoFitContents HTML Table" -rowArray $rowdata

	The 'rowArray' paramater is mandatory to build the table, but it is not set as such in the function - if nothing is passed, the table will be empty.

	Colors and Bold/Italics Flags are shown below:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack     

#>

Function FormatHTMLTable
{
	Param([string]$tableheader,
	[string]$tablewidth="auto",
	[string]$fontName="Calibri",
	[int]$fontSize=2,
	[switch]$noBorder=$false,
	[int]$noHeadCols=1,
	[object[]]$rowArray=@(),
	[object[]]$fixedWidth=@(),
	[object[]]$columnArray=@())

	$HTMLBody = "<b><font face='" + $fontname + "' size='" + ($fontsize + 1) + "'>" + $tableheader + "</font></b>"

	If($columnArray.Length -eq 0)
	{
		$NumCols = $noHeadCols + 1
	}  # means we have no column headers, just a table
	Else
	{
		$NumCols = $columnArray.Length
	}  # need to add one for the color attrib

	If($Null -ne $rowArray)
	{
		$NumRows = $rowArray.length + 1
	}
	Else
	{
		$NumRows = 1
	}

	If($noBorder)
	{
		$htmlbody += "<table border='0' width='" + $tablewidth + "'>"
	}
	Else
	{
		$htmlbody += "<table border='1' width='" + $tablewidth + "'>"
	}

	If(!($columnArray.Length -eq 0))
	{
		$htmlbody += "<tr>"

		For($columnIndex = 0; $columnIndex -lt $NumCols; $columnindex+=2)
		{
			$tmp = CheckHTMLColor $columnArray[$columnIndex+1]
			If($fixedWidth.Length -eq 0)
			{
				$htmlbody += "<td style=""background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}
			Else
			{
				$htmlbody += "<td style=""width:$($fixedWidth[$columnIndex/2]); background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}

			If($columnArray[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "<b>"
			}
			If($columnArray[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "<i>"
			}
			If($Null -ne $columnArray[$columnIndex])
			{
				If($columnArray[$columnIndex] -eq " " -or $columnArray[$columnIndex].length -eq 0)
				{
					$htmlbody += "&nbsp;&nbsp;&nbsp;"
				}
				Else
				{
					$found = $false
					For($i=0;$i -lt $columnArray[$columnIndex].length;$i+=2)
					{
						If($columnArray[$columnIndex][$i] -eq " ")
						{
							$htmlbody += "&nbsp;"
						}
						If($columnArray[$columnIndex][$i] -ne " ")
						{
							Break
						}
					}
					$htmlbody += $columnArray[$columnIndex]
				}
			}
			Else
			{
				$htmlbody += "&nbsp;&nbsp;&nbsp;"
			}
			If($columnArray[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "</b>"
			}
			If($columnArray[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "</i>"
			}
			$htmlbody += "</font></td>"
		}
		$htmlbody += "</tr>"
	}
	$rowindex = 2
	If($Null -ne $rowArray)
	{
		AddHTMLTable $fontName $fontSize -colCount $numCols -rowCount $NumRows -rowInfo $rowArray -fixedInfo $fixedWidth
		$rowArray = @()
		$htmlbody = "</table>"
	}
	Else
	{
		$HTMLBody += "</table>"
	}	
	out-file -FilePath $Script:FileName1 -Append -InputObject $HTMLBody 4>$Null 
}
#endregion

#region other HTML functions
#***********************************************************************************************************
# CheckHTMLColor - Called from AddHTMLTable WriteHTMLLine and FormatHTMLTable
#***********************************************************************************************************
Function CheckHTMLColor
{
	Param($hash)

	If($hash -band $htmlwhite)
	{
		Return $htmlwhitemask
	}
	If($hash -band $htmlred)
	{
		Return $htmlredmask
	}
	If($hash -band $htmlcyan)
	{
		Return $htmlcyanmask
	}
	If($hash -band $htmlblue)
	{
		Return $htmlbluemask
	}
	If($hash -band $htmldarkblue)
	{
		Return $htmldarkbluemask
	}
	If($hash -band $htmllightblue)
	{
		Return $htmllightbluemask
	}
	If($hash -band $htmlpurple)
	{
		Return $htmlpurplemask
	}
	If($hash -band $htmlyellow)
	{
		Return $htmlyellowmask
	}
	If($hash -band $htmllime)
	{
		Return $htmllimemask
	}
	If($hash -band $htmlmagenta)
	{
		Return $htmlmagentamask
	}
	If($hash -band $htmlsilver)
	{
		Return $htmlsilvermask
	}
	If($hash -band $htmlgray)
	{
		Return $htmlgraymask
	}
	If($hash -band $htmlblack)
	{
		Return $htmlblackmask
	}
	If($hash -band $htmlorange)
	{
		Return $htmlorangemask
	}
	If($hash -band $htmlmaroon)
	{
		Return $htmlmaroonmask
	}
	If($hash -band $htmlgreen)
	{
		Return $htmlgreenmask
	}
	If($hash -band $htmlolive)
	{
		Return $htmlolivemask
	}
}

Function SetupHTML
{
	Write-Verbose "$(Get-Date): Setting up HTML"
    If($AddDateTime)
    {
		$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).html"
    }

    $htmlhead = "<html><head><meta http-equiv='Content-Language' content='da'><title>" + $Script:Title + "</title></head><body>"
    #echo $htmlhead > $FileName1
	out-file -FilePath $Script:FileName1 -Force -InputObject $HTMLHead 4>$Null
}

Function HTMLHeatMap
{
    Param([decimal]$PValue)
    
    Switch($PValue)
    {
        {$_ -lt 70}{return $htmlgreen}
        {$_ -ge 70 -and $_ -lt 80}{return $htmlyellow}
        {$_ -ge 80 -and $_ -lt 90}{return $htmlorange}
        {$_ -ge 90 -and $_ -le 100}{return $htmlred}
    }

}

#endregion

#region Iain's Word table functions

<#
.Synopsis
	Add a table to a Microsoft Word document
.DESCRIPTION
	This function adds a table to a Microsoft Word document from either an array of
	Hashtables or an array of PSCustomObjects.

	Using this function is quicker than setting each table cell individually but can
	only utilise the built-in MS Word table autoformats. Individual tables cells can
	be altered after the table has been appended to the document (a table reference
	is returned).
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. Column headers will display the key names as defined.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -List

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. No column headers will be added, in a ListView format.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray

	This example adds table to the MS Word document, utilising all note property names
	the array of PSCustomObjects. Column headers will display the note property names.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -Columns FirstName,LastName,EmailAddress

	This example adds a table to the MS Word document, but only using the specified
	key names: FirstName, LastName and EmailAddress. If other keys are present in the
	array of Hashtables they will be ignored.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray -Columns FirstName,LastName,EmailAddress -Headers "First Name","Last Name","Email Address"

	This example adds a table to the MS Word document, but only using the specified
	PSCustomObject note properties: FirstName, LastName and EmailAddress. If other note
	properties are present in the array of PSCustomObjects they will be ignored. The
	display names for each specified column header has been overridden to display a
	custom header. Note: the order of the header names must match the specified columns.
#>

Function AddWordTable
{
	[CmdletBinding()]
	Param
	(
		# Array of Hashtable (including table headers)
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='Hashtable', Position=0)]
		[ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Hashtable,
		# Array of PSCustomObjects
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='CustomObject', Position=0)]
		[ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
		# Array of Hashtable key names or PSCustomObject property names to include, in display order.
		# If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Columns = $Null,
		# Array of custom table header strings in display order.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Headers = $Null,
		# AutoFit table behavior.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [int] $AutoFit = -1,
		# List view (no headers)
		[Switch] $List,
		# Grid lines
		[Switch] $NoGridLines,
		[Switch] $NoInternalGridLines,
		# Built-in Word table formatting style constant
		# Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
		[Parameter(ValueFromPipelineByPropertyName=$True)] [int] $Format = 0
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName);
		## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
		If(($Null -eq $Columns) -and ($Null -ne $Headers)) 
		{
			Write-Warning "No columns specified and therefore, specified headers will be ignored.";
			$Columns = $Null;
		}
		ElseIf(($Null -ne $Columns) -and ($Null -ne $Headers)) 
		{
			## Check if number of specified -Columns matches number of specified -Headers
			If($Columns.Length -ne $Headers.Length) 
			{
				Write-Error "The specified number of columns does not match the specified number of headers.";
			}
		} ## end elseif
	} ## end Begin

	Process
	{
		## Build the Word table data string to be converted to a range and then a table later.
		[System.Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder;

		Switch ($PSCmdlet.ParameterSetName) 
		{
			'CustomObject' 
			{
				If($Null -eq $Columns) 
				{
					## Build the available columns from all available PSCustomObject note properties
					[string[]] $Columns = @();
					## Add each NoteProperty name to the array
					ForEach($Property in ($CustomObject | Get-Member -MemberType NoteProperty)) 
					{ 
						$Columns += $Property.Name; 
					}
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{ 
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}

				## Iterate through each PSCustomObject
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Object in $CustomObject) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Object.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end foreach
				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
			} ## end CustomObject

			Default 
			{   ## Hashtable
				If($Null -eq $Columns) 
				{
					## Build the available columns from all available hashtable keys. Hopefully
					## all Hashtables have the same keys (they should for a table).
					$Columns = $Hashtable[0].Keys;
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{ 
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}
                
				## Iterate through each Hashtable
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Hash in $Hashtable) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Hash.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end foreach

				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f $Hashtable.Count);
			} ## end default
		} ## end switch

		## Create a MS Word range and set its text to our tab-delimited, concatenated string
		Write-Debug ("$(Get-Date): `t`tBuilding table range");
		$WordRange = $Script:Doc.Application.Selection.Range;
		$WordRange.Text = $WordRangeString.ToString();

		## Create hash table of named arguments to pass to the ConvertToTable method
		$ConvertToTableArguments = @{ Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByTabs; }

		## Negative built-in styles are not supported by the ConvertToTable method
		If($Format -ge 0) 
		{
			$ConvertToTableArguments.Add("Format", $Format);
			$ConvertToTableArguments.Add("ApplyBorders", $True);
			$ConvertToTableArguments.Add("ApplyShading", $True);
			$ConvertToTableArguments.Add("ApplyFont", $True);
			$ConvertToTableArguments.Add("ApplyColor", $True);
			If(!$List) 
			{ 
				$ConvertToTableArguments.Add("ApplyHeadingRows", $True); 
			}
			$ConvertToTableArguments.Add("ApplyLastRow", $True);
			$ConvertToTableArguments.Add("ApplyFirstColumn", $True);
			$ConvertToTableArguments.Add("ApplyLastColumn", $True);
		}

		## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
		## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
		Write-Debug ("$(Get-Date): `t`tConverting range to table");
		## Store the table reference just in case we need to set alternate row coloring
		$WordTable = $WordRange.GetType().InvokeMember(
			"ConvertToTable",                               # Method name
			[System.Reflection.BindingFlags]::InvokeMethod, # Flags
			$Null,                                          # Binder
			$WordRange,                                     # Target (self!)
			([Object[]]($ConvertToTableArguments.Values)),  ## Named argument values
			$Null,                                          # Modifiers
			$Null,                                          # Culture
			([String[]]($ConvertToTableArguments.Keys))     ## Named argument names
		);

		## Implement grid lines (will wipe out any existing formatting
		If($Format -lt 0) 
		{
			Write-Debug ("$(Get-Date): `t`tSetting table format");
			$WordTable.Style = $Format;
		}

		## Set the table autofit behavior
		If($AutoFit -ne -1) 
		{ 
			$WordTable.AutoFitBehavior($AutoFit); 
		}

		If(!$List)
		{
			#the next line causes the heading row to flow across page breaks
			$WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue;
		}

		If(!$NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleSingle;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}
		If($NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleNone;
		}
		If($NoInternalGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}

		Return $WordTable;

	} ## end Process
}

<#
.Synopsis
	Sets the format of one or more Word table cells
.DESCRIPTION
	This function sets the format of one or more table cells, either from a collection
	of Word COM object cell references, an individual Word COM object cell reference or
	a hashtable containing Row and Column information.

	The font name, font size, bold, italic , underline and shading values can be used.
.EXAMPLE
	SetWordCellFormat -Hashtable $Coordinates -Table $TableReference -Bold

	This example sets all text to bold that is contained within the $TableReference
	Word table, using an array of hashtables. Each hashtable contain a pair of co-
	ordinates that is used to select the required cells. Note: the hashtable must
	contain the .Row and .Column key names. For example:
	@ { Row = 7; Column = 3 } to set the cell at row 7 and column 3 to bold.
.EXAMPLE
	$RowCollection = $Table.Rows.First.Cells
	SetWordCellFormat -Collection $RowCollection -Bold -Size 10

	This example sets all text to size 8 and bold for all cells that are contained
	within the first row of the table.
	Note: the $Table.Rows.First.Cells returns a collection of Word COM cells objects
	that are in the first table row.
.EXAMPLE
	$ColumnCollection = $Table.Columns.Item(2).Cells
	SetWordCellFormat -Collection $ColumnCollection -BackgroundColor 255

	This example sets the background (shading) of all cells in the table's second
	column to red.
	Note: the $Table.Columns.Item(2).Cells returns a collection of Word COM cells objects
	that are in the table's second column.
.EXAMPLE
	SetWordCellFormat -Cell $Table.Cell(17,3) -Font "Tahoma" -Color 16711680

	This example sets the font to Tahoma and the text color to blue for the cell located
	in the table's 17th row and 3rd column.
	Note: the $Table.Cell(17,3) returns a single Word COM cells object.
#>

Function SetWordCellFormat 
{
	[CmdletBinding(DefaultParameterSetName='Collection')]
	Param (
		# Word COM object cell collection reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName='Collection', Position=0)] [ValidateNotNullOrEmpty()] $Collection,
		# Word COM object individual cell reference
		[Parameter(Mandatory=$true, ParameterSetName='Cell', Position=0)] [ValidateNotNullOrEmpty()] $Cell,
		# Hashtable of cell co-ordinates
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=0)] [ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Coordinates,
		# Word COM object table reference
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=1)] [ValidateNotNullOrEmpty()] $Table,
		# Font name
		[Parameter()] [AllowNull()] [string] $Font = $null,
		# Font color
		[Parameter()] [AllowNull()] $Color = $null,
		# Font size
		[Parameter()] [ValidateNotNullOrEmpty()] [int] $Size = 0,
		# Cell background color
		[Parameter()] [AllowNull()] $BackgroundColor = $null,
		# Force solid background color
		[Switch] $Solid,
		[Switch] $Bold,
		[Switch] $Italic,
		[Switch] $Underline
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'." -f $PSCmdlet.ParameterSetName);
	}

	Process 
	{
		Switch ($PSCmdlet.ParameterSetName) 
		{
			'Collection' {
				ForEach($Cell in $Collection) 
				{
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				} # end foreach
			} # end Collection
			'Cell' 
			{
				If($Bold) { $Cell.Range.Font.Bold = $true; }
				If($Italic) { $Cell.Range.Font.Italic = $true; }
				If($Underline) { $Cell.Range.Font.Underline = 1; }
				If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
				If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
				If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
				If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
				If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
			} # end Cell
			'Hashtable' 
			{
				ForEach($Coordinate in $Coordinates) 
				{
					$Cell = $Table.Cell($Coordinate.Row, $Coordinate.Column);
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				}
			} # end Hashtable
		} # end switch
	} # end process
}

<#
.Synopsis
	Sets alternate row colors in a Word table
.DESCRIPTION
	This function sets the format of alternate rows within a Word table using the
	specified $BackgroundColor. This function is expensive (in performance terms) as
	it recursively sets the format on alternate rows. It would be better to pick one
	of the predefined table formats (if one exists)? Obviously the more rows, the
	longer it takes :'(

	Note: this function is called by the AddWordTable function if an alternate row
	format is specified.
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 255

	This example sets every-other table (starting with the first) row and sets the
	background color to red (wdColorRed).
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 39423 -Seed Second

	This example sets every other table (starting with the second) row and sets the
	background color to light orange (weColorLightOrange).
#>

Function SetWordTableAlternateRowColor 
{
	[CmdletBinding()]
	Param (
		# Word COM object table reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)] [ValidateNotNullOrEmpty()] $Table,
		# Alternate row background color
		[Parameter(Mandatory=$true, Position=1)] [ValidateNotNull()] [int] $BackgroundColor,
		# Alternate row starting seed
		[Parameter(ValueFromPipelineByPropertyName=$true, Position=2)] [ValidateSet('First','Second')] [string] $Seed = 'First'
	)

	Process 
	{
		$StartDateTime = Get-Date;
		Write-Debug ("{0}: `t`tSetting alternate table row colors.." -f $StartDateTime);

		## Determine the row seed (only really need to check for 'Second' and default to 'First' otherwise
		If($Seed.ToLower() -eq 'second') 
		{ 
			$StartRowIndex = 2; 
		}
		Else 
		{ 
			$StartRowIndex = 1; 
		}

		For($AlternateRowIndex = $StartRowIndex; $AlternateRowIndex -lt $Table.Rows.Count; $AlternateRowIndex += 2) 
		{ 
			$Table.Rows.Item($AlternateRowIndex).Shading.BackgroundPatternColor = $BackgroundColor;
		}

		## I've put verbose calls in here we can see how expensive this functionality actually is.
		$EndDateTime = Get-Date;
		$ExecutionTime = New-TimeSpan -Start $StartDateTime -End $EndDateTime;
		Write-Debug ("{0}: `t`tDone setting alternate row style color in '{1}' seconds" -f $EndDateTime, $ExecutionTime.TotalSeconds);
	}
}
#endregion

#region general script functions
Function VISetup( [string] $VIServer )
{
    $script:startTime = Get-Date

    # Check for root
    # http://blogs.technet.com/b/heyscriptingguy/archive/2011/05/11/check-for-admin-credentials-in-a-powershell-script.aspx
    If(!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
    {
        Write-Host "`nThis script is not running as administrator - this is required to set global PowerCLI parameters. You may see PowerCLI warnings.`n"
        #Exit
    }

    Write-Verbose "$(Get-Date): Setting up VMware PowerCLI"
    #Check to see if PowerCLI is installed
    If($PCLICustom)
    {
        Write-Verbose "$(Get-Date): Custom PowerCLI Install location"
        $PCLIPath = "$(Select-FolderDialog)\Initialize-PowerCLIEnvironment.ps1" 4>$Null
    }
    ElseIf($env:PROCESSOR_ARCHITECTURE -like "*AMD64*"){$PCLIPath = "C:\Program Files (x86)\VMware\Infrastructure\vSphere PowerCLI\Scripts\Initialize-PowerCLIEnvironment.ps1"}
    Else{$PCLIPath = "C:\Program Files\VMware\Infrastructure\vSphere PowerCLI\Scripts\Initialize-PowerCLIEnvironment.ps1"}

    If (Test-Path $PCLIPath)
    {
            # grab the PWD before PCLI resets it to C:\
            $tempPWD = $pwd
            Import-Module $PCLIPath *>$Null
    }
    Else
    {
            Write-Host "`nPowerCLI does not appear to be installed - please install the latest version of PowerCLI. This script will now exit."
            Write-Host "*** If PowerCLI was installed to a non-Default location, please use the -PCLICustom parameter ***`n"
            Exit
        }

    $xPowerCLIVer = Get-PowerCLIVersion 4>$Null
    Write-Verbose "$(Get-Date): Loaded PowerCLI version $($xPowerCLIVer.Major).$($xPowerCLIVer.Minor)"
    If($xPowerCLIVer.Major -lt 5 -or ($xPowerCLIVer.Major -eq 5 -and $xPowerCLIVer.Minor -lt 1))
    {
        Write-Host "`nPowerCLI version $($xPowerCLIVer.Major).$($xPowerCLIVer.Minor) is installed. PowerCLI version 5.1 or later is required to run this script. `nPlease install the latest version and run this script again. This script will now exit."
        Exit
    }
    
    #Set PCLI defaults and reset PWD
    cd $tempPWD 4>$Null
    Write-Verbose "$(Get-Date): Setting PowerCLI global Configuration"
    Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -DisplayDeprecationWarnings $False -Confirm:$False *>$Null

    #Are we already connected to VC?
    If($global:DefaultVIServer)
    {
        Write-Host "`nIt appears PowerCLI is already connected to a VCenter Server. Please use the 'Disconnect-VIServer' cmdlet to disconnect any sessions before running inventory."
        Exit
    }

    #Connect to VI Server
    Write-Verbose "$(Get-Date): Connecting to VIServer: $($VIServer)"
    $Script:VCObj = Connect-VIServer $VIServer 4>$Null
    If($Export){$Script:VCObj | Export-Clixml .\Export\VCObj.xml 4>$Null}

    #Verify we successfully connected
    If(!($?))
    {
            Write-Host "Connecting to vCenter failed with the following error: $($Error[0].Exception.Message.substring($Error[0].Exception.Message.IndexOf("Connect-VIServer") + 16).Trim()) This script will now exit."
            Exit
        }

    #[string]$script:Title = "VMware Inventory Report - $VIServerName"
    #SetFileName1andFileName2 "$($VIServer)-Inventory"

}

Function Select-FolderDialog
{
    # http://stackoverflow.com/questions/11412617/get-a-folder-path-from-the-explorer-menu-to-a-powershell-variable
    param([string]$Description="Select PowerCLI Scripts Directory - Default is C:\Program Files (x86)\VMware\Infrastructure\vSphere PowerCLI\Scripts\",[string]$RootFolder="Desktop")

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null     

    $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
    $objForm.Rootfolder = $RootFolder
    $objForm.Description = $Description
    $Show = $objForm.ShowDialog()
    If ($Show -eq "OK")
    {
        Return $objForm.SelectedPath
    }
    Else
    {
        Write-Error "Operation cancelled by user."
    }
}

Function SetGlobals
{
    Write-Verbose "$(Get-Date): Gathering VMware data"
    ## Any Get used more than once is set to a global variable to limit the number of calls to PowerCLI
    ## Export commands from http://blogs.technet.com/b/heyscriptingguy/archive/2011/09/06/learn-how-to-save-powershell-objects-for-offline-analysis.aspx
    If($Export)
    {
        If(!(Test-Path .\Export))
        {
            New-Item .\Export -type directory *>$Null
        }
        Write-Verbose "$(Get-Date): Gathering Compute data"
        $Script:VMHosts = Get-VMHost 4>$Null| Sort Name 
        Get-Cluster 4>$Null| Sort Name | Export-Clixml .\Export\Cluster.xml 4>$Null
        $VMHosts | Sort Name | Export-Clixml .\Export\VMHost.xml 4>$Null
        Get-Datastore 4>$Null| Sort Name | Export-Clixml .\Export\Datastore.xml 4>$Null
        Get-Snapshot -VM * 4>$Null| Export-Clixml .\Export\Snapshot.xml 4>$Null
        Get-AdvancedSetting -Entity $VIServerName 4>$Null| Where {$_.Type -eq "VIServer" -and ($_.Name -like "mail.smtp.port" -or $_.Name -like "mail.smtp.server" -or $_.Name -like "mail.sender" -or $_.Name -like "VirtualCenter.FQDN")} | Export-Clixml .\Export\vCenterAdv.xml 4>$Null
        Get-AdvancedSetting -Entity ($VMHosts | Where {$_.ConnectionState -like "*Connected*" -or $_.ConnectionState -like "*Maintenance*"}).Name 4>$Null| Where {$_.Name -like "Syslog.global.logdir" -or $_.Name -like "Syslog.global.loghost"} | Export-Clixml .\Export\HostsAdv.xml 4>$Null
        Get-VMHostService -VMHost * 4>$Null| Export-Clixml .\Export\HostService.xml 4>$Null
        Write-Verbose "$(Get-Date): Gathering Virtual Machine data"
        $Script:VirtualMachines = Get-VM 4>$Null| Sort Name
        $VirtualMachines | Export-Clixml .\Export\VM.xml 4>$Null
        Get-ResourcePool 4>$Null| Sort Name | Export-Clixml .\Export\ResourcePool.xml 4>$Null
        Get-View 4>$Null(Get-View ServiceInstance 4>$Null).Content.PerfManager | Export-Clixml .\Export\vCenterStats.xml 4>$Null
        Get-View ServiceInstance 4>$Null| Export-Clixml .\Export\ServiceInstance.xml 4>$Null
        (((Get-View extensionmanager).ExtensionList).Description 4>$Null) | Export-Clixml .\Export\Plugins.xml 4>$Null
        Get-View 4>$Null(Get-View serviceInstance 4>$Null| Select -First 1).Content.LicenseManager | Export-Clixml .\Export\Licensing.xml 4>$Null
        Get-VIPermission 4>$Null| Sort Entity | Export-Clixml .\Export\VIPerms.xml 4>$Null
        Get-VIRole 4>$Null| Sort Name | Export-Clixml .\Export\VIRoles.xml 4>$Null
        BuildDRSRules | Export-Clixml .\Export\DRSRules.xml 4>$Null
        If($Full)
        {
            If(!(Test-Path .\Export\VMDetail))
            {
                New-Item .\Export\VMDetail -type directory *>$Null
            }      
            Write-Verbose "$(Get-Date): Gathering Networking data"
            Get-VMHostNetworkAdapter 4>$Null| Export-Clixml .\Export\HostNetwork.xml 4>$Null
            Get-VirtualSwitch 4>$Null| Export-Clixml .\Export\vSwitch.xml 4>$Null
            Get-NetworkAdapter * 4>$Null| Export-Clixml .\Export\NetworkAdapter.xml 4>$Null
            Get-VirtualPortGroup 4>$Null| Export-Clixml .\Export\PortGroup.xml 4>$Null
        }
    }
    ElseIf($Import)
    {
        $script:startTime = Get-Date
        ## Check for export directory
        If(Test-Path .\Export)
        {
            $Script:Clusters = Import-Clixml .\Export\Cluster.xml
            $Script:VMHosts = Import-Clixml .\Export\VMHost.xml
            $Script:Datastores = Import-Clixml .\Export\Datastore.xml
            $Script:Snapshots = Import-Clixml .\Export\Snapshot.xml
            $Script:HostAdvSettings = Import-Clixml .\Export\HostsAdv.xml
            $Script:VCAdvSettings = Import-Clixml .\Export\vCenterAdv.xml
            $Script:VCObj = Import-Clixml .\Export\VCObj.xml
            $Script:HostServices = Import-Clixml .\Export\HostService.xml
            $Script:VirtualMachines = Import-Clixml .\Export\VM.xml
            $Script:Resources = Import-Clixml .\Export\ResourcePool.xml
            $SCript:VMPlugins = Import-Clixml .\Export\Plugins.xml
            $Script:vCenterStatistics = Import-Clixml .\Export\vCenterStats.xml
            $Script:VCLicensing = Import-Clixml .\Export\Licensing.xml
            $Script:VIPerms = Import-Clixml .\Export\VIPerms.xml
            $Script:VIRoles = Import-Clixml .\Export\VIRoles.xml
            $Script:DRSRules = Import-Clixml .\Export\DRSRules.xml
            If(Test-Path .\Export\RegSQL.xml){$Script:RegSQL = Import-Clixml .\Export\RegSQL.xml}
            If($Full)
            {
                $Script:HostNetAdapters = Import-Clixml .\Export\HostNetwork.xml
                $Script:VirtualSwitches = Import-Clixml .\Export\vSwitch.xml
                $Script:VMNetworkAdapters = Import-Clixml .\Export\NetworkAdapter.xml
                $Script:VirtualPortGroups = Import-Clixml .\Export\PortGroup.xml
            }
            $Script:VIServerName = (($VCAdvSettings) | Where {$_.Name -like "VirtualCenter.FQDN"}).Value

        }
        Else
        {
            ## VMware Export not found, exit script
            Write-Host "Import option set, but no Export data directory found. Please copy the Export folder into the same folder as this script and run it again. This script will now exit."
            Exit
        }
    }
    ElseIf($Issues)
    {
        Write-Verbose "$(Get-Date): Gathering Compute data"
        $Script:Clusters = Get-Cluster 4>$Null | Sort Name
        $Script:VMHosts = Get-VMHost 4>$Null | Sort Name
        $Script:Datastores = Get-Datastore 4>$Null | Sort Name
        Write-Verbose "$(Get-Date): Gathering Virtual Machine data"
        $Script:VirtualMachines = Get-VM 4>$Null | Sort Name
        $Script:Snapshots = Get-Snapshot -VM * 4>$Null | Sort VM
    }
    Else
    {
        Write-Verbose "$(Get-Date): Gathering Compute data"
        $Script:Clusters = Get-Cluster 4>$Null| Sort Name 
        $Script:VMHosts = Get-VMHost 4>$Null| Sort Name 
        $Script:Datastores = Get-Datastore 4>$Null| Sort Name 
        $Script:HostAdvSettings = Get-AdvancedSetting -Entity ($VMHosts | Where {$_.ConnectionState -like "*Connected*" -or $_.ConnectionState -like "*Maintenance*"}).Name 4>$Null
        $Script:VCAdvSettings = Get-AdvancedSetting -Entity $VIServerName 4>$Null
        $Script:HostServices = Get-VMHostService -VMHost * 4>$Null
        Write-Verbose "$(Get-Date): Gathering Virtual Machine data"
        $Script:VirtualMachines = Get-VM 4>$Null| Sort Name 
        $Script:Resources = Get-ResourcePool 4>$Null| Sort Name 
        $Script:VMPlugins = (((Get-View extensionmanager 4>$Null).ExtensionList).Description) 
        $Script:ServiceInstance = Get-View ServiceInstance 4>$Null
        $Script:vCenterStatistics = Get-View ($ServiceInstance).Content.PerfManager 4>$Null
        $Script:VCLicensing = Get-View ($ServiceInstance | Select -First 1).Content.LicenseManager 4>$Null
        $Script:VIPerms = Get-VIPermission 4>$Null| Sort Entity 
        $Script:VIRoles = Get-VIRole 4>$Null| Sort Name 
        $Script:DRSRules = BuildDRSGroupsRules

        If($Full)
        {
            $Script:Snapshots = Get-Snapshot -VM * 4>$Null
            Write-Verbose "$(Get-Date): Gathering Networking data"
            $Script:HostNetAdapters = Get-VMHostNetworkAdapter 4>$Null
            $Script:VirtualSwitches = Get-VirtualSwitch 4>$Null
            $Script:VMNetworkAdapters = Get-NetworkAdapter * 4>$Null
            $Script:VirtualPortGroups = Get-VirtualPortGroup 4>$Null
        }
    }

    ## Cannot use Get-VMHostStorage in the event of a completely headless server with no local storage (cdrom, local disk, etc)
    ## Get-VMHostStorage with a wildcard will fail on headless servers so it must be called for each host
}

Function AddStatsChart
{
    Param
    (
        # Needed paramaters: ChartType, Title, Width, Length, ChartName, Format
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, ParameterSetName='StatData', Position=0)][ValidateNotNullOrEmpty()] $StatData,
        [Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] $StatData2 = $null,
        [Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] $StatData3 = $null,
        [Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [string] $Data1Label = $null,
        [Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [string] $Data2Label = $null,
        [Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [string] $Data3Label = $null,
        [Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [string] $Type = $null,
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)] [string] $Title = $null,
        [Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [int] $Width = $null,
        [Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [int] $Length = $null,
        [Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [string] $ExportName = $null,
        [Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [string] $ExportType = $null,
        [Parameter(ValueFromPipelineByPropertyName=$true)] [AllowNull()] [switch] $Legend = $false
    )

    Process
    {
        # load the appropriate assemblies 
        [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
        [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
        
        # http://blogs.technet.com/b/richard_macdonald/archive/2009/04/28/3231887.aspx
        # create the chart object
        $Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart 
        ## Ensure .NET Charting is installed
        If(!$?)
        {
            ## Assembly not loaded - exit
            Write-Host "`nMicrosoft Chart Controls for .NET is not installed but required to generate chart images. `nPlease install the latest version and run this script again. This script will now exit. `nhttp://www.microsoft.com/en-us/download/details.aspx?id=14422"
            Exit            
        }  
        $Chart.Width = $Width
        $Chart.Height = $Length

        # create a chartarea to draw on and add to chart 
        $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea 
        $Chart.ChartAreas.Add($ChartArea)

        # get Time\value data from get-stats
        $VMTime = @(ForEach($stamp in $StatData){$stamp.TimeStamp})
        $VMValue = @(ForEach($stamp2 in $StatData){$stamp2.Value})

        # convert KB memory to GB
        If(($StatData.unit | Select -Unique) -eq "KB")
        {
            $KBConversion = $true
            $ChartArea.AxisY.Title = "GB"
            $tempArr = @()
            ForEach($ValueData in $VMValue)
            {
                $tempArr += $ValueData / 1048576
            }
            $VMValue = $tempArr
        }
        Else
        {
            If($StatData.unit | Select -Unique){ $ChartArea.AxisY.Title = ($StatData.unit | Select -Unique) }
        }

        # add titles and data series
        [void]$Chart.Titles.Add($Title)
        

        If(!$Data1Label){$Data1Label = "Data 1"}
        [void]$Chart.Series.Add($Data1Label) 
        $Chart.Series[$Data1Label].Points.DataBindXY($VMTime, $VMValue)
        $Chart.Series[$Data1Label].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::$Type
        $Chart.Series[$Data1Label].Color = [System.Drawing.Color]::Blue

        # add 2nd data series
        If ($StatData2)
        {
            If(!$Data2Label){$Data2Label = "Data 2"}
            $VMTime2 = @(ForEach($stamp in $StatData2){$stamp.TimeStamp})
            $VMValue2 = @(ForEach($stamp2 in $StatData2){$stamp2.Value})
            If($KBConversion)
            {
                $tempArr = @()
                ForEach($ValueData2 in $VMValue2)
                {
                    $tempArr += $ValueData2 / 1048576
                }
                $VMValue2 = $tempArr
            }
            [void]$Chart.Series.Add($Data2Label)
            $Chart.Series[$Data2Label].Points.DataBindXY($VMTime2, $VMValue2)
            $Chart.Series[$Data2Label].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::$Type
            $Chart.Series[$Data2Label].Color = [System.Drawing.Color]::Red
        }
        
        # add 3rd data series
        If ($StatData3)
        {
            If(!$Data3Label){$Data3Label = "Data 3"}
            $VMTime3 = @(ForEach($stamp in $StatData3){$stamp.TimeStamp})
            $VMValue3 = @(ForEach($stamp2 in $StatData3){$stamp2.Value})
            If($KBConversion)
            {
                $tempArr = @()
                ForEach($ValueData3 in $VMValue3)
                {
                    $tempArr += $ValueData3 / 1048576
                }
                $VMValue3 = $tempArr
            }
            [void]$Chart.Series.Add($Data3Label)
            $Chart.Series[$Data3Label].Points.DataBindXY($VMTime3, $VMValue3)
            $Chart.Series[$Data3Label].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::$Type
            $Chart.Series[$Data3Label].Color = [System.Drawing.Color]::Green
        }

        $Chart.BackColor = [System.Drawing.Color]::Transparent

        # add a legend
        If($Legend)
        {
            $ChartLegend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend
            $Chart.Legends.Add($Legend) | Out-Null
        }

        If(!$ExportName)
        {
            $Chart.SaveImage(".\Chart$($Title).png", "PNG")
            $Script:Word.Selection.InlineShapes.AddPicture("$($PSScriptRoot)\Chart$($Title).png") | Out-Null
            Remove-Item "$($PSScriptRoot)\Chart$($Title).png" -Force
        }

    } # End Process
}

Function BuildDRSGroupsRules
{
    ## From http://www.vnugglets.com/2011/07/backupexport-full-drs-rule-info-via.html
    Get-View -ViewType ClusterComputeResource -Property Name, ConfigurationEx 4>$Null| %{
        ## if the cluster has any DRS rules
        if ($_.ConfigurationEx.Rule -ne $null) {
            $viewCurrClus = $_
            $DRSGroupsRules = @()
            $viewCurrClus.ConfigurationEx.Rule | %{
                $oRuleInfo = New-Object -Type PSObject -Property @{
                    ClusterName = $viewCurrClus.Name
                    RuleName = $_.Name
                    RuleType = $_.GetType().Name
                    bRuleEnabled = $_.Enabled
                    bMandatory = $_.Mandatory
                } 
 
                ## add members to the output object, to be populated in a bit
                "bKeepTogether,VMNames,VMGroupName,VMGroupMembers,AffineHostGrpName,AffineHostGrpMembers,AntiAffineHostGrpName,AntiAffineHostGrpMembers".Split(",") | %{Add-Member -InputObject $oRuleInfo -MemberType NoteProperty -Name $_ -Value $null}
 
                ## switch statement based on the object type of the .NET view object
                switch ($_){
                    ## if it is a ClusterVmHostRuleInfo rule, get the VM info from the cluster View object
                    #   a ClusterVmHostRuleInfo item "identifies virtual machines and host groups that determine virtual machine placement"
                    {$_ -is [VMware.Vim.ClusterVmHostRuleInfo]} {
                        $oRuleInfo.VMGroupName = $_.VmGroupName
                        ## get the VM group members' names
                        $oRuleInfo.VMGroupMembers = (Get-View -Property Name -Id ($viewCurrClus.ConfigurationEx.Group | ?{($_ -is [VMware.Vim.ClusterVmGroup]) -and ($_.Name -eq $oRuleInfo.VMGroupName)}).Vm 4>$Null| %{$_.Name}) -join ","
                        $oRuleInfo.AffineHostGrpName = $_.AffineHostGroupName
                        ## get affine hosts' names
                        $oRuleInfo.AffineHostGrpMembers = if ($_.AffineHostGroupName -ne $null) {(Get-View -Property Name -Id ($viewCurrClus.ConfigurationEx.Group | ?{($_ -is [VMware.Vim.ClusterHostGroup]) -and ($_.Name -eq $oRuleInfo.AffineHostGrpName)}).Host 4>$Null| %{$_.Name}) -join ","}
                        $oRuleInfo.AntiAffineHostGrpName = $_.AntiAffineHostGroupName
                        ## get anti-affine hosts' names
                        $oRuleInfo.AntiAffineHostGrpMembers = if ($_.AntiAffineHostGroupName -ne $null) {(Get-View -Property Name -Id ($viewCurrClus.ConfigurationEx.Group | ?{($_ -is [VMware.Vim.ClusterHostGroup]) -and ($_.Name -eq $oRuleInfo.AntiAffineHostGrpName)}).Host 4>$Null| %{$_.Name}) -join ","}
                        break;
                    } 
                    ## if ClusterAffinityRuleSpec (or AntiAffinity), get the VM names (using Get-View)
                    {($_ -is [VMware.Vim.ClusterAffinityRuleSpec]) -or ($_ -is [VMware.Vim.ClusterAntiAffinityRuleSpec])} {
                        $oRuleInfo.VMNames = if ($_.Vm.Count -gt 0) {(Get-View -Property Name -Id $_.Vm 4>$Null| %{$_.Name}) -join ","}
                    } 
                    {$_ -is [VMware.Vim.ClusterAffinityRuleSpec]} {
                        $oRuleInfo.bKeepTogether = $true
                    } 
                    {$_ -is [VMware.Vim.ClusterAntiAffinityRuleSpec]} {
                        $oRuleInfo.bKeepTogether = $false
                    }
                    default {"none of the above"}
                }
 
                $DRSGroupsRules += $oRuleInfo
            } 
        } 
    } 
    Return $DRSGroupsRules
}

Function truncate
{
    Param([string]$strIn, [int]$length)
    If($strIn.Length -gt $length){ "$($strIn.Substring(0, $length))..." }
    Else{ $strIn }
}

Function ShowScriptOptions
{
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): AddDateTime   : $($AddDateTime)"
	Write-Verbose "$(Get-Date): Chart         : $($Chart)"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): Company Name  : $($Script:CoName)"
	}
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): Cover Page    : $($CoverPage)"
	}
	Write-Verbose "$(Get-Date): Dev           : $($Dev)"
	If($Dev)
	{
		Write-Verbose "$(Get-Date): DevErrorFile  : $($Script:DevErrorFile)"
	}
	Write-Verbose "$(Get-Date): Export        : $($Export)"
	Write-Verbose "$(Get-Date): Filename1     : $($Script:filename1)"
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Filename2     : $($Script:filename2)"
	}
	Write-Verbose "$(Get-Date): Folder        : $($Folder)"
	Write-Verbose "$(Get-Date): From          : $($From)"
	Write-Verbose "$(Get-Date): Full          : $($Full)"
	Write-Verbose "$(Get-Date): Import        : $($Import)"
	Write-Verbose "$(Get-Date): Issues        : $($Issues)"
	Write-Verbose "$(Get-Date): Save As HTML  : $($HTML)"
	Write-Verbose "$(Get-Date): Save As PDF   : $($PDF)"
	Write-Verbose "$(Get-Date): Save As TEXT  : $($TEXT)"
	Write-Verbose "$(Get-Date): Save As WORD  : $($MSWORD)"
	Write-Verbose "$(Get-Date): ScriptInfo    : $($ScriptInfo)"
	Write-Verbose "$(Get-Date): Smtp Port     : $($SmtpPort)"
	Write-Verbose "$(Get-Date): Smtp Server   : $($SmtpServer)"
	Write-Verbose "$(Get-Date): Title         : $($Script:Title)"
	Write-Verbose "$(Get-Date): To            : $($To)"
	Write-Verbose "$(Get-Date): Use SSL       : $($UseSSL)"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): User Name     : $($UserName)"
	}
	Write-Verbose "$(Get-Date): VIServerName  : $($VIServerName)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): OS Detected   : $($Script:RunningOS)"
	Write-Verbose "$(Get-Date): PoSH version  : $($Host.Version)"
	Write-Verbose "$(Get-Date): PSCulture     : $($PSCulture)"
	Write-Verbose "$(Get-Date): PSUICulture   : $($PSUICulture)"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): Word language : $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date): Word version  : $($Script:WordProduct)"
	}
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Script start  : $($Script:StartTime)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
}

Function validStateProp( [object] $object, [string] $topLevel, [string] $secondLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	if( $object )
	{
		If( ( gm -Name $topLevel -InputObject $object ) )
		{
			If( ( gm -Name $secondLevel -InputObject $object.$topLevel ) )
			{
				Return $True
			}
		}
	}
	Return $False
}

Function SaveandCloseDocumentandShutdownWord
{
	#bug fix 1-Apr-2014
	#reset Grammar and Spelling options back to their original settings
	$Script:Word.Options.CheckGrammarAsYouType = $Script:CurrentGrammarOption
	$Script:Word.Options.CheckSpellingAsYouType = $Script:CurrentSpellingOption

	Write-Verbose "$(Get-Date): Save and Close document and Shutdown Word"
	If($Script:WordVersion -eq $wdWord2010)
	{
		#the $saveFormat below passes StrictMode 2
		#I found this at the following two links
		#http://blogs.technet.com/b/bshukla/archive/2011/09/27/3347395.aspx
		#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Saving DOCX file"
		}
		If($AddDateTime)
		{
			$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
			If($PDF)
			{
				$Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
			}
		}
		Write-Verbose "$(Get-Date): Running $Script:WordProduct and detected operating system $($Script:RunningOS)"
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$Script:Doc.SaveAs([REF]$Script:FileName1, [ref]$SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$saveFormat)
		}
	}
	ElseIf($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
	{
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Saving DOCX file"
		}
		If($AddDateTime)
		{
			$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
			If($PDF)
			{
				$Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
			}
		}
		Write-Verbose "$(Get-Date): Running $Script:WordProduct and detected operating system $($Script:RunningOS)"
		$Script:Doc.SaveAs2([REF]$Script:FileName1, [ref]$wdFormatDocumentDefault)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$wdFormatPDF)
		}
	}

	Write-Verbose "$(Get-Date): Closing Word"
	$Script:Doc.Close()
	Write-Verbose "$(Get-Date): Waiting 10 seconds to allow Word to save file"
	Start-Sleep -Seconds 10
	$Script:Word.Quit()
	Write-Verbose "$(Get-Date): Waiting 10 seconds to allow Word to fully close"
	Start-Sleep -Seconds 10
	If($PDF)
	{
		[int]$cnt = 0
		While(Test-Path $Script:FileName1)
		{
			$cnt++
			If($cnt -gt 1)
			{
				Write-Verbose "$(Get-Date): Waiting another 10 seconds to allow Word to fully close (try # $($cnt))"
				Start-Sleep -Seconds 10
				$Script:Word.Quit()
				If($cnt -gt 2)
				{
					#kill the winword process

					#find out our session (usually "1" except on TS/RDC or Citrix)
					$SessionID = (Get-Process -PID $PID).SessionId
					
					#Find out if winword is running in our session
					$wordprocess = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}).Id
					If($wordprocess -gt 0)
					{
						Write-Verbose "$(Get-Date): Attempting to stop WinWord process # $($wordprocess)"
						Stop-Process $wordprocess -EA 0
					}
				}
			}
			Write-Verbose "$(Get-Date): Attempting to delete $($Script:FileName1) since only $($Script:FileName2) is needed (try # $($cnt))"
			Remove-Item $Script:FileName1 -EA 0 4>$Null
		}
	}
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	If(Test-Path variable:global:word)
	{
		Remove-Variable -Name word -Scope Global 4>$Null
	}
	$SaveFormat = $Null
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	
	#is the winword process still running? kill it

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId

	#Find out if winword is running in our session
	$wordprocess = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}).Id
	If($wordprocess -gt 0)
	{
		Write-Verbose "$(Get-Date): WinWord process is still running. Attempting to stop WinWord process # $($wordprocess)"
		Stop-Process $wordprocess -EA 0
	}
}

Function SaveandCloseTextDocument
{
	If($AddDateTime)
	{
		$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	}

	Write-Output $Global:Output | Out-File $Script:Filename1 4>$Null
}

Function SaveandCloseHTMLDocument
{
	Out-File -FilePath $Script:FileName1 -Append -InputObject "<p></p></body></html>" 4>$Null
}

Function SetFileName1andFileName2
{
	Param([string]$OutputFileName)
	
	If($Folder -eq "")
	{
		$pwdpath = $pwd.Path
	}
	Else
	{
		$pwdpath = $Folder
	}

	If($pwdpath.EndsWith("\"))
	{
		#remove the trailing \
		$pwdpath = $pwdpath.SubString(0, ($pwdpath.Length - 1))
	}

	#set $filename1 and $filename2 with no file extension
	If($AddDateTime)
	{
		[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName)"
		If($PDF)
		{
			[string]$Script:FileName2 = "$($pwdpath)\$($OutputFileName)"
		}
	}

	If($MSWord -or $PDF)
	{
		CheckWordPreReq

		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).docx"
			If($PDF)
			{
				[string]$Script:FileName2 = "$($pwdpath)\$($OutputFileName).pdf"
			}
		}

		SetupWord
	}
	ElseIf($Text)
	{
		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).txt"
		}
		ShowScriptOptions
	}
	ElseIf($HTML)
	{
		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).html"
		}
		SetupHTML
		ShowScriptOptions
	}
}

Function TestComputerName
{
	Param([string]$Cname)
	If(![String]::IsNullOrEmpty($CName)) 
	{
		#get computer name
		#first test to make sure the computer is reachable
		Write-Verbose "$(Get-Date): Testing to see if $($CName) is online and reachable"
		If(Test-Connection -ComputerName $CName -quiet)
		{
			Write-Verbose "$(Get-Date): Server $($CName) is online."
		}
		Else
		{
			Write-Verbose "$(Get-Date): Computer $($CName) is offline"
			$ErrorActionPreference = $SaveEAPreference
			Write-Error "`n`n`t`tComputer $($CName) is offline.`nScript cannot continue.`n`n"
			Exit
		}
	}

	#if computer name is localhost, get actual computer name
	If($CName -eq "localhost")
	{
		$CName = $env:ComputerName
		Write-Verbose "$(Get-Date): Computer name has been renamed from localhost to $($CName)"
		Return $CName
	}

	#if computer name is an IP address, get host name from DNS
	#http://blogs.technet.com/b/gary/archive/2009/08/29/resolve-ip-addresses-to-hostname-using-powershell.aspx
	#help from Michael B. Smith
	$ip = $CName -as [System.Net.IpAddress]
	If($ip)
	{
		$Result = [System.Net.Dns]::gethostentry($ip)
		
		If($? -and $Null -ne $Result)
		{
			$CName = $Result.HostName
			Write-Verbose "$(Get-Date): Computer name has been renamed from $($ip) to $($CName)"
			Return $CName
		}
		Else
		{
			Write-Warning "Unable to resolve $($CName) to a hostname"
		}
	}
	Else
	{
		#computer is online but for some reason $ComputerName cannot be converted to a System.Net.IpAddress
	}
	Return $CName
}

Function ProcessDocumentOutput
{
	If($MSWORD -or $PDF)
	{
		SaveandCloseDocumentandShutdownWord
	}
	ElseIf($Text)
	{
		SaveandCloseTextDocument
	}
	ElseIf($HTML)
	{
		SaveandCloseHTMLDocument
	}

	$GotFile = $False

	If($PDF)
	{
		If(Test-Path "$($Script:FileName2)")
		{
			Write-Verbose "$(Get-Date): $($Script:FileName2) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName2)"
			Write-Error "Unable to save the output file, $($Script:FileName2)"
		}
	}
	Else
	{
		If(Test-Path "$($Script:FileName1)")
		{
			Write-Verbose "$(Get-Date): $($Script:FileName1) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName1)"
			Write-Error "Unable to save the output file, $($Script:FileName1)"
		}
	}
	
	#email output file if requested
	If($GotFile -and ![System.String]::IsNullOrEmpty( $SmtpServer ))
	{
		If($PDF)
		{
			$emailAttachment = $Script:FileName2
		}
		Else
		{
			$emailAttachment = $Script:FileName1
		}
		SendEmail $emailAttachment
	}
}

Function AbortScript
{
	If($MSWord -or $PDF)
	{
		$Script:Word.quit()
		Write-Verbose "$(Get-Date): System Cleanup"
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
		If(Test-Path variable:global:word)
		{
			Remove-Variable -Name word -Scope Global
		}
	}
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	Write-Verbose "$(Get-Date): Script has been aborted"
	$ErrorActionPreference = $SaveEAPreference
	Exit
}
#endregion

#region Summary and vCenter functions
Function ProcessSummary
{
    Write-Verbose "$(Get-Date): Processing Summary page"
    If($MSWord -or $PDF)
    {
        $Selection.InsertNewPage()
        WriteWordLine 1 0 "vCenter Summary"

        $TableRange = $doc.Application.Selection.Range
        $Table = $doc.Tables.Add($TableRange, 1, 5)
	    $Table.Style = $Script:MyHash.Word_TableGrid
        $Table.Cell(1,1).Range.Text = "Legend: "
        $Table.Cell(1,2).Range.Text = "0% - 69%"
        $Table.Cell(1,3).Range.Text = "70% - 79%"
        $Table.Cell(1,4).Range.Text = "80% - 89%"
        $Table.Cell(1,5).Range.Text = "90% - 100%"

        SetWordCellFormat -Cell $Table.Cell(1,2) -BackgroundColor 7405514
        SetWordCellFormat -Cell $Table.Cell(1,3) -BackgroundColor 9434879
        SetWordCellFormat -Cell $Table.Cell(1,4) -BackgroundColor 42495
        SetWordCellFormat -Cell $Table.Cell(1,5) -BackgroundColor 238

        $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
	    $Table.AutoFitBehavior($wdAutoFitContent)

	    #return focus back to document
	    $doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

	    #move to the end of the current document
	    $selection.EndKey($wdStory,$wdMove) | Out-Null
	    $TableRange = $Null
	    $Table = $Null
        WriteWordLine 0 0 ""

    }
    ElseIF($HTML)
    {
        WriteHTMLLine 1 0 "vCenter Summary"
        $rowData = @()
        $columnHeaders = @("Legend",$htmlwhite,'0% - 69%',$htmlgreen,'70% - 79%',$htmlyellow,'80% - 89%',$htmlorange,'90% - 100%',$htmlred)
        FormatHTMLTable "" -columnArray $columnHeaders
        WriteHTMLLine 0 1 ""
    }
    ElseIf($Text)
    {
        Line 0 "vCenter Summary"
    }

    ## Cluster Summary
    If($MSWord -or $PDF)
    {
        WriteWordLine 2 0 "Cluster Summary"

        ## Create an array of hashtables
	    [System.Collections.Hashtable[]] $ClusterWordTable = @();
	    ## Seed the row index from the second row
	    [int] $CurrentServiceIndex = 2;

        ForEach($Cluster in $Clusters)
        {
            ## Add the required key/values to the hashtable
	        $WordTableRowHash = @{ 
	        Cluster = $Cluster.Name;
	        HostsinCluster = @($VMHosts | Where {$_.IsStandAlone -eq $False -and $_.Parent -like $Cluster.Name}).Count;
	        HAEnabled = $Cluster.HAEnabled;
	        DRSEnabled = $Cluster.DrsEnabled;
            DRSAutomation = $Cluster.DrsAutomationLevel;
            VMCount = @($VirtualMachines | Where {$_.VMHost -in @($VMHosts | Where {$_.ParentId -like $Cluster.Id}).Name}).Count;
	        }
	        ## Add the hash to the array
	        $ClusterWordTable += $WordTableRowHash;
	        $CurrentServiceIndex++;
    
        }

        ## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
	    $Table = AddWordTable -Hashtable $ClusterWordTable `
	    -Columns Cluster, HostsinCluster, HAEnabled, DRSEnabled, DRSAutomation, VMCount `
	    -Headers "Cluster Name", "Host Count", "HA Enabled", "DRS Enabled", "DRS Automation Level", "VM Count" `
	    -Format $wdTableGrid `
	    -AutoFit $wdAutoFitContent;

	    ## IB - Set the header row format
	    SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

	    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

	    FindWordDocumentEnd
	    $Table = $Null

	    WriteWordLine 0 0 ""
    }
    ElseIf($HTML)
    {
        $rowData = @()
        $columnHeaders = @("Cluster Name",($htmlsilver -bor $htmlbold),"Host Count",($htmlsilver -bor $htmlbold),"HA Enabled",($htmlsilver -bor $htmlbold),"DRS Enabled",($htmlsilver -bor $htmlbold),"DRS Automation Level",($htmlsilver -bor $htmlbold),"VM Count",($htmlsilver -bor $htmlbold))

        ForEach($Cluster in $Clusters)
        {
            $rowData += @(,($Cluster.Name,$htmlwhite,@($VMHosts | Where {$_.IsStandAlone -eq $False -and $_.Parent -like $Cluster.Name}).Count,$htmlwhite,$Cluster.HAEnabled,$htmlwhite,$Cluster.DrsEnabled,$htmlwhite,$Cluster.DrsAutomationLevel,$htmlwhite,@($VirtualMachines | Where {$_.VMHost -in @($VMHosts | Where {$_.ParentId -like $Cluster.Id}).Name}).Count,$htmlwhite))
        }
        FormatHTMLTable "Cluster Summary" -rowArray $rowData -columnArray $columnHeaders
        WriteHTMLLine 0 1 ""
    }

    ##Host summary

    If($MSWord -or $PDF)
    {
        WriteWordLine 2 0 "Host Summary"
        ## Create an array of hashtables
	    [System.Collections.Hashtable[]] $HostWordTable = @();
	    ## Seed the row index from the second row
	    [int] $CurrentServiceIndex = 2;

        $heatMap = @{Row = @(); Column = @(); Color = @()}
        ForEach($VMHost in $VMHosts)
        {
            If ($VMHost.IsStandAlone)
            {
                $xStandAlone = "Standalone"
            }
            Else
            {
                $xStandAlone = $VMhost.Parent
            }

            ## Add the required key/values to the hashtable
	        $WordTableRowHash = @{ 
	        VMHost = $VMHost.Name;
	        ConnectionState = $VMHost.ConnectionState;
            ESXVersion = $VMHost.Version;
	        ClusterMember = $xStandAlone;
	        CPUPercent = $("{0:P2}" -f $($VMHost.CpuUsageMhz / $VMHost.CpuTotalMhz));
            MemoryPercent = $("{0:P2}" -f $($VMhost.MemoryUsageGB / $VMHost.MemoryTotalGB));
            VMCount = @(($VirtualMachines) | Where {$_.VMHost -like $VMHost.Name}).Count;
	        }

            ## Build VMHost heatmap
            Switch([decimal]($WordTableRowHash.CPUPercent -replace '%'))
            {
                {$_ -lt 70}{$heatMap.Row += @($CurrentServiceIndex); $heatMap.Column += @(5); $heatMap.Color += @(7405514);}
                {$_ -ge 70 -and $_ -lt 80}{$heatMap.Row += @($CurrentServiceIndex); $heatMap.Column += @(5); $heatMap.Color += @(9434879);}
                {$_ -ge 80 -and $_ -lt 90}{$heatMap.Row += @($CurrentServiceIndex); $heatMap.Column += @(5); $heatMap.Color += @(42495);}
                {$_ -ge 90 -and $_ -le 100}{$heatMap.Row += @($CurrentServiceIndex); $heatMap.Column += @(5); $heatMap.Color += @(238);}
            }
            Switch([decimal]($WordTableRowHash.MemoryPercent -replace '%'))
            {
                {$_ -lt 70}{$heatMap.Row += @($CurrentServiceIndex); $heatMap.Column += @(6); $heatMap.Color += @(7405514);}
                {$_ -ge 70 -and $_ -lt 80}{$heatMap.Row += @($CurrentServiceIndex); $heatMap.Column += @(6); $heatMap.Color += @(9434879);}
                {$_ -ge 80 -and $_ -lt 90}{$heatMap.Row += @($CurrentServiceIndex); $heatMap.Column += @(6); $heatMap.Color += @(42495);}
                {$_ -ge 90 -and $_ -le 100}{$heatMap.Row += @($CurrentServiceIndex); $heatMap.Column += @(6); $heatMap.Color += @(238);}
            }
	        ## Add the hash to the array
	        $HostWordTable += $WordTableRowHash;
	        $CurrentServiceIndex++;
    
        }

        ## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
	    $Table = AddWordTable -Hashtable $HostWordTable `
	    -Columns VMHost, ConnectionState, ESXVersion, ClusterMember, CPUPercent, MemoryPercent, VMCount `
	    -Headers "Host Name", "Connection State", "ESX Version", "Parent Cluster", "CPU Used %", "Memory Used %", "VM Count" `
	    -Format $wdTableGrid `
	    -AutoFit $wdAutoFitContent;

	    ## IB - Set the header row format
        SetWordTableAlternateRowColor $Table $wdColorGray05 "Second"
	    SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
        For($i = 0; $i -lt $heatMap.Row.Count; $i++)
        {
            SetWordCellFormat -Cell $Table.Cell($heatMap.Row[$i],$heatMap.Column[$i]) -BackgroundColor $heatMap.Color[$i]
        }

	    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

	    FindWordDocumentEnd
	    $Table = $Null

	    WriteWordLine 0 0 ""
    }
    ElseIf($HTML)
    {
        $rowData = @()
        $columnHeaders = @("Host Name",($htmlsilver -bor $htmlbold),"Connection State",($htmlsilver -bor $htmlbold),"ESX Version",($htmlsilver -bor $htmlbold),"Parent Cluster",($htmlsilver -bor $htmlbold),'CPU Used %',($htmlsilver -bor $htmlbold),'Memory Used %',($htmlsilver -bor $htmlbold),"VM Count",($htmlsilver -bor $htmlbold))
        ForEach($VMHost in $VMHosts)
        {
            If ($VMHost.IsStandAlone)
            {
                $xStandAlone = "Standalone"
            }
            Else
            {
                $xStandAlone = $VMhost.Parent
            }

            $cpuPerc = HTMLHeatMap (($VMHost.CpuUsageMhz / $VMHost.CpuTotalMhz) * 100)
            $memPerc = HTMLHeatMap (($VMhost.MemoryUsageGB / $VMHost.MemoryTotalGB) * 100)
            $rowData += @(,($VMHost.Name,$htmlwhite,$VMHost.ConnectionState,$htmlwhite,$VMHost.Version,$htmlwhite,$xStandAlone,$htmlwhite,$("{0:P2}" -f $($VMHost.CpuUsageMhz / $VMHost.CpuTotalMhz)),$cpuPerc,$("{0:P2}" -f $($VMhost.MemoryUsageGB / $VMHost.MemoryTotalGB)),$memPerc,@(($VirtualMachines) | Where {$_.VMHost -like $VMHost.Name}).Count,$htmlwhite))
        }

        FormatHTMLTable "Host Summary" -rowArray $rowData -columnArray $columnHeaders
        WriteHTMLLine 0 1 ""
    }

    ##Datastore summary
    If($MSWord -or $PDF)
    {
        WriteWordLine 2 0 "Datastore Summary"
        ## Create an array of hashtables
	    [System.Collections.Hashtable[]] $DatastoreWordTable = @();
	    ## Seed the row index from the second row
	    [int] $CurrentServiceIndex = 2;

        $heatMap = @{Row = @(); Column = @(); Color = @()}
        ForEach($Datastore in $Datastores)
        {
            $WordTableRowHash = @{
            Datastore = $Datastore.Name;
            DSType = $Datastore.Type;
            DSTotalCap = $("{0:N2}" -f $Datastore.CapacityGB + " GB");
            DSFreeCap = $("{0:N2}" -f $Datastore.FreeSpaceGB + " GB");
            DSFreePerc = $("{0:P2}" -f $(($Datastore.CapacityGB - $Datastore.FreeSpaceGB) / $Datastore.CapacityGB));
            }
    	
            ## Build Datastore summary heatmap
            Switch([decimal]($WordTableRowHash.DSFreePerc -replace '%'))
            {
                {$_ -lt 70}{$heatMap.Row += @($CurrentServiceIndex); $heatMap.Column += @(5); $heatMap.Color += @(7405514);}
                {$_ -ge 70 -and $_ -lt 80}{$heatMap.Row += @($CurrentServiceIndex); $heatMap.Column += @(5); $heatMap.Color += @(9434879);}
                {$_ -ge 80 -and $_ -lt 90}{$heatMap.Row += @($CurrentServiceIndex); $heatMap.Column += @(5); $heatMap.Color += @(42495);}
                {$_ -ge 90 -and $_ -le 100}{$heatMap.Row += @($CurrentServiceIndex); $heatMap.Column += @(5); $heatMap.Color += @(238);}
            }
            ## Add the hash to the array
	        $DatastoreWordTable += $WordTableRowHash;
	        $CurrentServiceIndex++;
        }

        ## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
	    $Table = AddWordTable -Hashtable $DatastoreWordTable `
	    -Columns Datastore, DSType, DSTotalCap, DSFreeCap, DSFreePerc `
	    -Headers "Datastore Name", "Type", "Total Capacity", "Free Space", "Percent Used" `
	    -Format $wdTableGrid `
	    -AutoFit $wdAutoFitContent;

	    ## IB - Set the header row format
        SetWordTableAlternateRowColor $Table $wdColorGray05 "Second"
	    SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
        For($i = 0; $i -lt $heatMap.Row.Count; $i++)
                {
        SetWordCellFormat -Cell $Table.Cell($heatMap.Row[$i],$heatMap.Column[$i]) -BackgroundColor $heatMap.Color[$i]
    }

	    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

	    FindWordDocumentEnd
	    $Table = $Null

	    WriteWordLine 0 0 ""
    }
    If($HTML)
    {
        $rowData = @()
        $columnHeaders = @("Datastore Name",($htmlsilver -bor $htmlbold),"Type",($htmlsilver -bor $htmlbold),"Total Capacity",($htmlsilver -bor $htmlbold),"Free Space",($htmlsilver -bor $htmlbold),"Percent Used",($htmlsilver -bor $htmlbold))

        ForEach($Datastore in $Datastores)
        {
            $dsPerc = HTMLHeatMap ((($Datastore.CapacityGB - $Datastore.FreeSpaceGB) / $Datastore.CapacityGB) * 100)
            $rowData += @(,($Datastore.Name,$htmlwhite,$Datastore.Type,$htmlwhite,$("{0:N2}" -f $Datastore.CapacityGB + " GB"),$htmlwhite,$("{0:N2}" -f $Datastore.FreeSpaceGB + " GB"),$htmlwhite,$("{0:P2}" -f $(($Datastore.CapacityGB - $Datastore.FreeSpaceGB) / $Datastore.CapacityGB)),$dsPerc))
        }

        FormatHTMLTable "Datastore Summary" -rowArray $rowData -columnArray $columnHeaders
    }
}

Function ProcessvCenter
{
    Write-Verbose "$(Get-Date): Processing vCenter Global Settings"
    If($MSWord -or $PDF)
	{
        $Selection.InsertNewPage()
		WriteWordLine 1 0 "vCenter Server"
	}
    ElseIf($HTML)
    {
        WriteHTMLLine 1 0 "vCenter Server"
    }
	ElseIf($Text)
	{
		Line 0 "vCenter Server"
	} 
    
    ## Global vCenter settings
    ## Try to get vCenter DSN if Windows Server
    If(!$Import)
    {
        $RemReg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $VIServerName)
        If($?)
        {
            $VCDBKey = "SOFTWARE\\VMware, Inc.\\VMware VirtualCenter\\DB"
            $VCDB = ($RemReg[0].OpenSubKey($VCDBKey,$true)).GetValue("1")

            $DBDetails = "SOFTWARE\\ODBC\\ODBC.INI\\$($VCDB)"   
        }
    }

    If($Export)
    {
        $RegExpObj = New-Object psobject
        $RegExpObj | Add-Member -Name VCDB -MemberType NoteProperty -Value $VCDB
        $RegExpObj | Add-Member -Name SQLDB -MemberType NoteProperty -Value ($RemReg[0].OpenSubKey($DBDetails,$true)).GetValue("Database")
        $RegExpObj | Add-Member -Name SQLServer -MemberType NoteProperty -Value ($RemReg[0].OpenSubKey($DBDetails,$true)).GetValue("Server")
        $RegExpObj | Add-Member -Name SQLUser -MemberType NoteProperty -Value ($RemReg[0].OpenSubKey($DBDetails,$true)).GetValue("LastUser")
        $RegExpObj | Export-Clixml .\Export\RegSQL.xml 4>$Null
    }
    If($MSWord -or $PDF)
    {
        WriteWordLine 2 0 "Global Settings"
        [System.Collections.Hashtable[]] $ScriptInformation = @()
        $ScriptInformation += @{ Data = "Name"; Value = (($VCAdvSettings) | Where {$_.Name -like "VirtualCenter.FQDN"}).Value;}
        $ScriptInformation += @{ Data = "Version"; Value = $VCObj.Version; }
        $ScriptInformation += @{ Data = "Build"; Value = $VCObj.Build; }
        If($Import -and $RegSQL)
        {
            $ScriptInformation += @{ Data = "DSN Name"; Value = $RegSQL.VCDB; }
            $ScriptInformation += @{ Data = "SQL Database"; Value = $RegSQL.SQLDB; }
            $ScriptInformation += @{ Data = "SQL Server"; Value = $RegSQL.SQLServer; }
            $ScriptInformation += @{ Data = "Last SQL User"; Value = $RegSQL.SQLUser; }
        }
        ElseIf($VCDB)
        {
            $ScriptInformation += @{ Data = "DSN Name"; Value = $VCDB; }
            $ScriptInformation += @{ Data = "SQL Database"; Value = ($RemReg[0].OpenSubKey($DBDetails,$true)).GetValue("Database"); }
            $ScriptInformation += @{ Data = "SQL Server"; Value = ($RemReg[0].OpenSubKey($DBDetails,$true)).GetValue("Server"); }
            $ScriptInformation += @{ Data = "Last SQL User"; Value = ($RemReg[0].OpenSubKey($DBDetails,$true)).GetValue("LastUser"); }
        }
        $ScriptInformation += @{ Data = "EMail Sender"; Value = (($VCAdvSettings) | Where {$_.Name -like "mail.sender"}).Value; }
        $ScriptInformation += @{ Data = "SMTP Server"; Value = (($VCAdvSettings) | Where {$_.Name -like "mail.smtp.server"}).Value; }
        $ScriptInformation += @{ Data = "SMTP Server Port"; Value = (($VCAdvSettings) | Where {$_.Name -like "mail.smtp.port"}).Value; }

        $Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 205;
		$Table.Columns.Item(2).Width = 200;

		#$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
    }
    ElseIf($HTML)
    {
        $rowdata = @()
        $colWidths = @("150px","200px")
        $rowdata += @(,("Server Name",($htmlsilver -bor $htmlbold),(($VCAdvSettings) | Where {$_.Name -like "VirtualCenter.FQDN"}).Value,$htmlwhite))
        $rowdata += @(,("Version",($htmlsilver -bor $htmlbold),$VCObj.Version,$htmlwhite))
        $rowdata += @(,("Build",($htmlsilver -bor $htmlbold),$VCObj.Build,$htmlwhite))
        If($Import -and $RegSQL)
        {
            $rowdata += @(,("DSN Name",($htmlsilver -bor $htmlbold),$RegSQL.VCDB,$htmlwhite))
            $rowdata += @(,("SQL Database",($htmlsilver -bor $htmlbold),$RegSQL.SQLDB,$htmlwhite))
            $rowdata += @(,("SQL Server",($htmlsilver -bor $htmlbold),$RegSQL.SQLServer,$htmlwhite))
            $rowdata += @(,("Last SQL User",($htmlsilver -bor $htmlbold),$RegSQL.SQLUser,$htmlwhite))         
        }
        ElseIf($VCDB)
        {
            $rowdata += @(,("DSN Name",($htmlsilver -bor $htmlbold),$VCDB,$htmlwhite))
            $rowdata += @(,("SQL Database",($htmlsilver -bor $htmlbold),($RemReg[0].OpenSubKey($DBDetails,$true)).GetValue("Database"),$htmlwhite))
            $rowdata += @(,("SQL Server",($htmlsilver -bor $htmlbold),($RemReg[0].OpenSubKey($DBDetails,$true)).GetValue("Server"),$htmlwhite))
            $rowdata += @(,("Last SQL User",($htmlsilver -bor $htmlbold),($RemReg[0].OpenSubKey($DBDetails,$true)).GetValue("LastUser"),$htmlwhite))
        }
        $rowdata += @(,("Email Sender",($htmlsilver -bor $htmlbold),(($VCAdvSettings) | Where {$_.Name -like "mail.sender"}).Value,$htmlwhite))
        $rowdata += @(,("SMTP Server",($htmlsilver -bor $htmlbold),(($VCAdvSettings) | Where {$_.Name -like "mail.smtp.server"}).Value,$htmlwhite))
        $rowdata += @(,("SMTP Server Port",($htmlsilver -bor $htmlbold),(($VCAdvSettings) | Where {$_.Name -like "mail.smtp.port"}).Value,$htmlwhite))

        FormatHTMLTable "General Settings" -noHeadCols 2 -rowArray $rowdata -fixedWidth $colWidths -tablewidth "350"
        WriteHTMLLine 0 1 ""
    }
    ElseIf($Text)
    {
        Line 0 "Global Settings" 
        Line 1 ""
        Line 1 "Server Name:`t`t" (($VCAdvSettings) | Where {$_.Name -like "VirtualCenter.FQDN"}).Value
        Line 1 "Version:`t`t" $VCObj.Version
        Line 1 "Build:`t`t`t" $VCObj.Build
        If($Import -and $RegSQL)
        {
            Line 1 "DSN Name:`t`t" $RegSQL.VCDB
            Line 1 "SQL Database:`t`t" $RegSQL.SQLDB
            Line 1 "SQL Server:`t`t" $RegSQL.SQLServer
            Line 1 "Last SQL User:`t`t" $RegSQL.SQLUser           
        }
        ElseIf($VCDB)
        {
            Line 1 "DSN Name:`t`t" $VCDB
            Line 1 "SQL Database:`t`t" ($RemReg[0].OpenSubKey($DBDetails,$true)).GetValue("Database")
            Line 1 "SQL Server:`t`t" ($RemReg[0].OpenSubKey($DBDetails,$true)).GetValue("Server")
            Line 1 "Last SQL User:`t`t" ($RemReg[0].OpenSubKey($DBDetails,$true)).GetValue("LastUser")
        }
        Line 1 "Email Sender:`t`t" (($VCAdvSettings) | Where {$_.Name -like "mail.sender"}).Value
        Line 1 "SMTP Server:`t`t" (($VCAdvSettings) | Where {$_.Name -like "mail.smtp.server"}).Value
        Line 1 "SMTP Server Port:`t" (($VCAdvSettings) | Where {$_.Name -like "mail.smtp.port"}).Value
        Line 0 ""
    }

    ## vCenter historical statistics
    $vCenterStats = @()
    If($MSWord -or $PDF)
    {
        ## Create an array of hashtables
	    [System.Collections.Hashtable[]] $vCenterStats = @();
	    ## Seed the row index from the second row
	    [int] $CurrentServiceIndex = 2;
        WriteWordLine 2 0 "Historical Statistics"
        Foreach($xStatLevel in $vCenterStatistics.HistoricalInterval)
        {
            Switch($xStatLevel.SamplingPeriod)
            {
                300{$xInterval = "5 Minutes"; Break}
                1800{$xInterval = "30 Minutes"; Break}
                7200{$xInterval = "2 Hours"; Break}
                86400{$xInterval = "1 Day"; Break}
            }
            ## Add the required key/values to the hashtable
	        $WordTableRowHash = @{
            IntervalDuration = $xInterval;
            IntervalEnabled = $xStatLevel.Enabled;
            SaveDuration = $xStatLevel.Name;
            StatsLevel = $xStatLevel.Level;
            }
            ## Add the hash to the array
	        $vCenterStats += $WordTableRowHash;
	        $CurrentServiceIndex++
        }

        ## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
	    $Table = AddWordTable -Hashtable $vCenterStats `
	    -Columns IntervalDuration, IntervalEnabled, SaveDuration, StatsLevel `
	    -Headers "Interval Duration", "Enabled", "Save For", "Statistics Level" `
	    -Format $wdTableGrid `
	    -AutoFit $wdAutoFitContent;

	    ## IB - Set the header row format
	    SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

	    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

	    FindWordDocumentEnd
	    $Table = $Null

	    WriteWordLine 0 0 ""
    }
    ElseIf($Text)
    {
        Line 0 "Historical Statistics" 
        Line 1 ""
        Line 1 "Interval Duration`tEnabled`t`tSave For`tStatistics Level"
        Foreach($xStatLevel in $vCenterStatistics.HistoricalInterval)
        {
            Switch($xStatLevel.SamplingPeriod)
            {
                300{$xInterval = "5 Min."; Break}
                1800{$xInterval = "30 Min."; Break}
                7200{$xInterval = "2 Hours"; Break}
                86400{$xInterval = "1 Day"; Break}
            }
            Line 1 "$($xInterval)`t`t`t$($xStatLevel.Enabled)`t`t$($xStatLevel.Name)`t$($xStatLevel.Level)"
        }
        Line 0 ""
    }
    ElseIf($HTML)
    {
        $rowdata = @()
        $columnHeaders = @("Interval Duration",($htmlsilver -bor $htmlbold),"Enabled",($htmlsilver -bor $htmlbold),"Save For",($htmlsilver -bor $htmlbold),"Statistics Level",($htmlsilver -bor $htmlbold))
        Foreach($xStatLevel in $vCenterStatistics.HistoricalInterval)
        {
            Switch($xStatLevel.SamplingPeriod)
            {
                300{$xInterval = "5 Min."; Break}
                1800{$xInterval = "30 Min."; Break}
                7200{$xInterval = "2 Hours"; Break}
                86400{$xInterval = "1 Day"; Break}
            }
            $rowdata += @(,($xInterval,$htmlwhite,$xStatLevel.Enabled,$htmlWhite,$xStatLevel.Name,$htmlWhite,$xStatLevel.Level,$htmlWhite))
        }
        FormatHTMLTable "Historical Statistics" -rowArray $rowdata -columnArray $columnHeaders
        WriteHTMLLine 0 0 " "
    }

    ## vCenter Licensing
    $vSphereLicInfo = @() 
    If($MSWord -or $PDF)
    {
        ## Create an array of hashtables
	    [System.Collections.Hashtable[]] $LicenseWordTable = @();
	    ## Seed the row index from the second row
	    [int] $CurrentServiceIndex = 2;
        WriteWordLine 2 0 "Licensing"
        #http://blogs.vmware.com/PowerCLI/2012/05/retrieving-license-keys-from-multiple-vCenters.html
        Foreach ($LicenseMan in $VCLicensing) 
        { 
            Foreach ($License in ($LicenseMan | Select -ExpandProperty Licenses))
            {
                ## Add the required key/values to the hashtable
	            $WordTableRowHash = @{ 
	            LicenseName = $License.Name;
	            LicenseKey = "*****" + $License.LicenseKey.Substring(23);
	            LicenseTotal = $License.Total;
	            LicenseUsed = $License.Used;
	            }
	            ## Add the hash to the array
	            $LicenseWordTable += $WordTableRowHash;
	            $CurrentServiceIndex++


            }
        }

        ## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
	    $Table = AddWordTable -Hashtable $LicenseWordTable `
	    -Columns LicenseName, LicenseKey, LicenseTotal, LicenseUsed `
	    -Headers "License Name", "Key Last 5", "Total Licenses", "Licenses Used" `
	    -Format $wdTableGrid `
	    -AutoFit $wdAutoFitContent;

	    ## IB - Set the header row format
	    SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

	    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

	    FindWordDocumentEnd
	    $Table = $Null

	    WriteWordLine 0 0 ""
    }
    ElseIf($Text)
    {
        Line 0 "Licensing" 
        Line 1 ""
        Line 1 "License Name`t`t`tKey Last 5`tTotal Licenses`tLicenses Used"
        Foreach ($LicenseMan in $VCLicensing) 
        { 
            Foreach ($License in ($LicenseMan | Select -ExpandProperty Licenses))
            {
                Line 1 "$($License.Name)`t*****$($License.LicenseKey.Substring(23))`t$($License.Total)`t`t$($License.Used)"
            }
        }
        Line 0 ""
    }
    ElseIf($HTML)
    {
        $rowdata = @()
        $columnHeaders = @("License Name",($htmlsilver -bor $htmlbold),"Key Last 5",($htmlsilver -bor $htmlbold),"Total Licenses",($htmlsilver -bor $htmlbold),"Licenses Used",($htmlsilver -bor $htmlbold))
        Foreach ($LicenseMan in $VCLicensing) 
        { 
            Foreach ($License in ($LicenseMan | Select -ExpandProperty Licenses))
            {
                $rowdata += @(,($License.Name,$htmlwhite,"*****$($License.LicenseKey.Substring(23))",$htmlwhite,$License.Total,$htmlwhite,$License.Used,$htmlwhite))
            }
        }
        FormatHTMLTable "Licensing" -rowArray $rowdata -columnArray $columnHeaders
        WriteHTMLLine 0 0 " "
    }

    ## vCenter Permissions
    If($MSWord -or $PDF)
    {
        ## Create an array of hashtables
	    [System.Collections.Hashtable[]] $PermsWordTable = @();
	    ## Seed the row index from the second row
	    [int] $CurrentServiceIndex = 2;
        WriteWordLine 2 0 "vCenter Permissions"
        foreach ($VIPerm in $VIPerms)
        {
            ## Add the required key/values to the hashtable
	        $WordTableRowHash = @{
            Entity = $VIPerm.Entity;
            Principal = $VIPerm.Principal;
            Role = $VIPerm.Role;
            }
            ## Add the hash to the array
            $PermsWordTable += $WordTableRowHash
            $CurrentServiceIndex++
        }

        ## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
	    $Table = AddWordTable -Hashtable $PermsWordTable `
	    -Columns Entity, Principal, Role `
	    -Headers "Entity", "Principal", "Role" `
	    -Format $wdTableGrid `
	    -AutoFit $wdAutoFitContent; 

	    ## IB - Set the header row format
	    SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

	    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

	    FindWordDocumentEnd
	    $Table = $Null

	    WriteWordLine 0 0 ""  

    }
    ElseIf($Text)
    {

    }
    ElseIf($HTML)
    {
        $rowData = @()
        $columnHeaders = @("Entity",($htmlsilver -bor $htmlbold),"Principal",($htmlsilver -bor $htmlbold),"Role",($htmlsilver -bor $htmlbold))
        foreach($VIPerm in $VIPerms)
        {
            $rowData += @(,($VIPerm.Entity,$htmlwhite,$VIPerm.Principal,$htmlwhite,$VIPerm.Role,$htmlwhite))
        }
        FormatHTMLTable "vCenter Permissions" -columnArray $columnHeaders -rowArray $rowdata
        WriteHTMLLine 0 0 " "
    }

    ## vCenter Role Perms
    If($MSWord -or $PDF)
    {
        WriteWordLine 2 0 "Active non-Standard vCenter Roles"
        foreach($role in ($VIPerms | select Role -Unique))
        {
            foreach($expandRole in $VIRoles | Where {$_.Name -eq $role.Role -and $_.IsSystem -eq $false})
            {
                WriteWordLine 0 0 $expandRole.Name -boldface $true
                foreach($privRole in $expandRole.PrivilegeList){WriteWordLine 0 0 $privRole -fontSize 8}
                WriteWordLine 0 0 ""
            }
        }
    }
    ElseIf($Text)
    {

    }
    ElseIf($HTML)
    {
        WriteHTMLLine 1 0 "Active non-Standard vCenter Roles"
        foreach($role in ($VIPerms | Select Role -Unique))
        {
            foreach($expandRole in $VIRoles | Where {$_.Name -eq $role.Role -and $_.IsSystem -eq $false})
            {
                WriteHTMLLine 0 0 $expandRole.Name -options $htmlBold -fontSize 3
                foreach($privRole in $expandRole.PrivilegeList){WriteHTMLLine 0 0 $privRole -fontSize 2}
                WriteHTMLLine 0 0 " "
            }
        }
    }

    ## vCenter Plugins
    $vSpherePlugins = @()
    If($MSWord -or $PDF)
    {
        ## Create an array of hashtables
	    [System.Collections.Hashtable[]] $PluginsWordTable = @();
	    ## Seed the row index from the second row
	    [int] $CurrentServiceIndex = 2;
        WriteWordLine 2 0 "vCenter Plugins"
        Foreach ($VMPlugin in $VMPlugins)
        {
            ## Add the required key/values to the hashtable
	        $WordTableRowHash = @{
            PluginName = $VMPlugin.Label;
            PluginDesc = $VMPlugin.Summary;
            }  
	        ## Add the hash to the array
	        $PluginsWordTable += $WordTableRowHash;
	        $CurrentServiceIndex++                       
        }  

        ## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
	    $Table = AddWordTable -Hashtable $PluginsWordTable `
	    -Columns PluginName, PluginDesc `
	    -Headers "Plugin", "Description" `
	    -Format $wdTableGrid `
	    -AutoFit $wdAutoFitContent;    

	    ## IB - Set the header row format
	    SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

	    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

	    FindWordDocumentEnd
	    $Table = $Null

	    WriteWordLine 0 0 ""    
              
    }
    ElseIf($Text)
    {
        Line 0 "Plugins"
        Line 1 ""
        LIne 1 "Plugin`t`t`tDescription"
        ForEach($VMPlugin in $VMPlugins)
        {
            Line 1 "$($VMPlugin.Label)`t$($VMPlugin.Summary)"
        }
    }
    ElseIf($HTML)
    {
        $rowdata = @()
        $columnHeaders = @("Plugin",($htmlsilver -bor $htmlbold),"Description",($htmlsilver -bor $htmlbold))
        Foreach($VMPlugin in $VMPlugins)
        {
            $rowData += @(,($VMPlugin.Label,$htmlwhite,$VMPlugin.Summary,$htmlwhite))
        }
        FormatHTMLTable "vCenter Plugins" -columnArray $columnHeaders -rowArray $rowdata
    }

}
#endregion

#region Hosts and Clusters functions
Function ProcessVMHosts
{
    Write-Verbose "$(Get-Date): Processing VMware Hosts"
    If($MSWord -or $PDF)
	{
        $Selection.InsertNewPage()
		WriteWordLine 1 0 "Hosts"
	}
	ElseIf($Text)
	{
		Line 0 "Hosts"
	}
    ElseIf($HTML)
    {
        WriteHTMLLine 1 0 "Hosts"
    }    

    If($? -and ($VMHosts))
    {
        $First = 0
        ForEach($VMHost in $VMHosts)
        {
            If($First -ne 0){$Selection.InsertNewPage()}
            OutputVMHosts $VMHost
            $First++
        }
    }
    ElseIf($? -and ($VMHosts -eq $Null))
    {
        Write-Warning "There are no ESX Hosts"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There are no ESX Hosts"
		}
		ElseIf($Text)
		{
			Line 1 "There are no ESX Hosts"
		}
		ElseIf($HTML)
		{
            WriteHTMLLine 0 1 "There are no ESX Hosts"
		}
    }
    Else
    {
    	Write-Warning "Unable to retrieve ESX Hosts"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "Unable to retrieve ESX Hosts"
		}
		ElseIf($Text)
		{
			Line 1 "Unable to retrieve ESX Hosts"
		}
		ElseIf($HTML)
		{
            WriteHTMLLine 0 1 "Unable to retrieve ESX Hosts"
		}
    }
}

Function OutputVMHosts
{
    Param([object] $VMHost)
    $xHostService = ($HostServices) | Where {$_.VMHostId -eq $VMHost.Id}
    If($Export) {(Get-VMHostNtpServer -VMHost $VMHost.Name 4>$Null) -join ", " | Export-Clixml .\Export\$($VMHost.Name)-NTP.xml 4>$Null} 
    $xHostAdvanced = ($HostAdvSettings) | Where {$_.Entity -like $VMHost.Name}
    ## Set vmhoststorage variable - will fail if host has no devices (headless - boot from USB\SD using NFS only)
    If($VMHost.PowerState -eq "PoweredOn")
    {
        If($Export)
        {
            Get-VMHostStorage -VMHost $VMHost 4>$Null| Where {$_.ScsiLun.LunType -notlike "*cdrom*"} | Export-Clixml .\Export\$($VMHost.Name)-Storage.xml 4>$Null
            If(!$?)
            {
                    Write-Warning ""
                    Write-Warning "Get-VMHostStorage failed. If $($VMHost) has no local storage and uses NFS only, the above error can be ignored."
                    Write-Warning ""
            }
        }
        ElseIf($Import)
        {
            $xVMHostStorage = Import-Clixml .\Export\$($VMHost.Name)-Storage.xml
        }
        Else
        {
            $xVMHostStorage = Get-VMHostStorage -VMHost $VMHost 4>$Null| Where {$_.ScsiLun.LunType -notlike "*cdrom*"} 4>$null
            If(!$?)
            {
                    Write-Warning ""
                    Write-Warning "Get-VMHostStorage failed. If $($VMHost) has no local storage and uses NFS only, the above error can be ignored."
                    Write-Warning ""
            }
        }
    }

    If ($VMHost.IsStandAlone)
    {
        $xStandAlone = "Standalone Host"
    }
    Else
    {
        $xStandAlone = "Clustered Host"
    }

    If ($VMHost.HyperthreadingActive)
    {
        $xHyperThreading = "Active"
    }
    Else
    {
        $xHyperThreading = "Disabled"
    }
    If ((($xHostService) | Where {$_.Key -eq "TSM-SSH"}).Running)
    {
        $xSSHService = "Running"
    }
    Else
    {
        $xSSHService = "Stopped"
    }
    If ((($xHostService) | Where {$_.Key -eq "ntpd"}).Running)
    {
        $xNTPService = "Running"
        If($Import)
        {
            $xNTPServers = Import-Clixml .\Export\$($VMHost.Name)-NTP.xml
        }
        Else
        {
            $xNTPServers = (Get-VMHostNtpServer -VMHost $VMHost.Name 4>$Null) -join ", " 4>$Null
        }
    }
    Else
    {
        $xNTPService = "Stopped"
    }
    If($xVMHostStorage.SoftwareIScsiEnabled)
    {
        $xiSCSI = "Enabled"
    }
    Else
    {
        $xiSCSI = "Disabled"
    }
    If($MSWord -or $PDF)
    {
        
        WriteWordLine 2 0 "Host: $($VMHost.Name)"
        [System.Collections.Hashtable[]] $ScriptInformation = @()
        $ScriptInformation += @{ Data = "Name"; Value = $VMHost.Name; }
        $ScriptInformation += @{ Data = "ESXi Version"; Value = $VMHost.Version; }
        $ScriptInformation += @{ Data = "ESXi Build"; Value = $VMHost.Build; }
		$ScriptInformation += @{ Data = "Power State"; Value = $VMHost.PowerState; }
		$ScriptInformation += @{ Data = "Connection State"; Value = $VMHost.ConnectionState; }
        $ScriptInformation += @{ Data = "Host Status"; Value = $xStandAlone; }
        If (!$VMHost.IsStandAlone)
        {
            $ScriptInformation += @{ Data = "Parent Object"; Value = $VMHost.Parent; }
        }
        If($VMHost.VMSwapfileDatastore)
        {
            $ScriptInformation += @{ Data = "VM Swapfile Datastore"; Value = $VMHost.VMSwapfileDatastore.Name; }
        }
        $ScriptInformation += @{ Data = "Manufacturer"; Value = $VMHost.Manufacturer; }
        $ScriptInformation += @{ Data = "Model"; Value = $VMHost.Model; }
        $ScriptInformation += @{ Data = "CPU Type"; Value = $VMHost.ProcessorType; }
        $ScriptInformation += @{ Data = "Maximum EVC Mode"; Value = $VMHost.MaxEVCMode; }
        $ScriptInformation += @{ Data = "CPU Core Count"; Value = $VMHost.NumCpu; }
        $ScriptInformation += @{ Data = "Hyperthreading"; Value = $xHyperthreading; }
        $ScriptInformation += @{ Data = "CPU Power Policy"; Value = $VMHost.ExtensionData.config.PowerSystemInfo.CurrentPolicy.ShortName; }
        $ScriptInformation += @{ Data = "Total Memory"; Value = "$([decimal]::round($VMHost.MemoryTotalGB)) GB"; }
        If($VMHost.PowerState -like "PoweredOn")
        {
            $ScriptInformation += @{ Data = "SSH Service Policy"; Value = (($xHostService) | Where {$_.Key -eq "TSM-SSH"}).Policy; }
            $ScriptInformation += @{ Data = "SSH Service Status"; Value = $xSSHService; }
            $ScriptInformation += @{ Data = "Scratch Log location"; Value = (($xHostAdvanced) | Where {$_.Name -eq "Syslog.global.logdir"}).Value; }
            $ScriptInformation += @{ Data = "Scratch Log remote host"; Value =  (($xHostAdvanced) | Where {$_.Name -eq "Syslog.global.loghost"}).Value; }
            $ScriptInformation += @{ Data = "NTP Service Policy"; Value = (($xHostService) | Where {$_.Key -eq "ntpd"}).Policy; }
            $ScriptInformation += @{ Data = "NTP Service Status"; Value = $xNTPService; }
            $ScriptInformation += @{ Data = "NFS Max Queue Depth"; Value = (($xHostAdvanced) | Where {$_.Name -eq "NFS.MaxQueueDepth"}).Value; }
            $ScriptInformation += @{ Data = "NFS Max Volumes"; Value = (($xHostAdvanced) | Where {$_.Name -eq "NFS.MaxVolumes"}).Value; }
            $ScriptInformation += @{ Data = "TCP IP Heap Size"; Value = (($xHostAdvanced) | Where {$_.Name -eq "Net.TcpipHeapSize"}).Value; }
            $ScriptInformation += @{ Data = "TCP IP Heap Max"; Value = (($xHostAdvanced) | Where {$_.Name -eq "Net.TcpipHeapMax"}).Value; }
            If($xVMHostStorage){$ScriptInformation += @{ Data = "Software iSCSI Service"; Value = $xiSCSI; }}
            If ((($xHostService) | Where {$_.Key -eq "ntpd"}).Running)
            {
                $ScriptInformation += @{ Data = "NTP Servers"; Value = $xNTPServers; }
            }
        }

        $Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 205;
		$Table.Columns.Item(2).Width = 200;

		# $Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
        If($xVMHostStorage)
        {
            WriteWordLine 0 0 "Block Storage"

            ## Create an array of hashtables
            [System.Collections.Hashtable[]] $ClusterWordTable = @();
            ## Seed the row index from the second row
            [int] $CurrentServiceIndex = 2;

            ForEach($xLUN in ($xVMHostStorage.ScsiLun) | Where {$_.LunType -notlike "*cdrom*"})
            {
                ## Add the required key/values to the hashtable
                $WordTableRowHash = @{
                Model = $xLUN.Model;
                Vendor = $xLUN.Vendor;
                Capacity = $("{0:N2}" -f $xLUN.CapacityGB + " GB");
                RuntimeName = $xLUN.RuntimeName;
                MultiPath = $xLUN.MultipathPolicy;
                Identifier =  truncate $xLUN.CanonicalName 16
                }
                ## Add the hash to the array
	            $ClusterWordTable += $WordTableRowHash;
	            $CurrentServiceIndex++;

            }

            ## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
	        $Table = AddWordTable -Hashtable $ClusterWordTable `
	        -Columns Model, Vendor, Capacity, RunTimeName, MultiPath, Identifier `
	        -Headers "Model", "Vendor", "Capacity", "Runtime Name", "MultiPath", "Identifier" `
	        -Format $wdTableGrid `
	        -AutoFit $wdAutoFitContent;

	        ## IB - Set the header row format
	        SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

	        $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

	        FindWordDocumentEnd
	        $Table = $Null

	        WriteWordLine 0 0 ""
        }

        If($VMHost.ConnectionState -like "*NotResponding*" -or $VMHost.PowerState -eq "PoweredOff")
        {
            WriteWordLine 0 0 "Note: $($VMHost.Name) is not responding or is in an unknown state - data in this and other reports will be missing or inaccurate." -italics $True
            WriteWordLine 0 0 ""
        }

        If($VMHost.PowerState -like "PoweredOn" -and $Chart)
        {
            $VMHostCPU = Get-Stat -Entity $VMHost.Name -Stat cpu.usage.average -Start (Get-Date).AddDays(-7) -Finish (Get-Date) -IntervalMins 30 | Where {$_.Instance -like ""}
            AddStatsChart -StatData $VMHostCPU -Title "$($VMHost.Name) CPU" -Width 275 -Length 200 -Type "Line"

            $VMHostGrant = Get-Stat -Entity $VMHost.Name -Stat mem.granted.average -Start (Get-Date).AddDays(-7) -Finish (Get-Date) -IntervalMins 30
            $VMHostActive = Get-Stat -Entity $VMHost.Name -Stat mem.active.average -Start (Get-Date).AddDays(-7) -Finish (Get-Date) -IntervalMins 30
            $VMHostBalloon = Get-Stat -Entity $VMHost.Name -Stat mem.vmmemctl.average -Start (Get-Date).AddDays(-7) -Finish (Get-Date) -IntervalMins 30
            AddStatsChart -StatData $VMHostGrant -StatData2 $VMHostActive -StatData3 $VMHostBalloon -Title "$($VMHost.Name) Memory" -Width 325 -Length 200 -Data1Label "Granted" -Data2Label "Active" -Data3Label "Balloon" -Legend -Type "Line"
            WriteWordLine 0 0 ""

            # Disk IO chart here...get-stats for NFS datastores may not be possible?

            $VMHostNetRec = get-stat -Entity $VMHost.Name -Stat "net.received.average" -Start (Get-Date).AddDays(-7) -Finish (Get-Date) -IntervalMins 30 | Where {$_.Instance -like ""}
            $VMHostNetTrans = get-stat -Entity $VMHost.Name -Stat "net.transmitted.average" -Start (Get-Date).AddDays(-7) -Finish (Get-Date) -IntervalMins 30 | Where {$_.Instance -like ""}
            AddStatsChart -StatData $VMHostNetRec -StatData2 $VMHostNetTrans -Title "$($VMHost.Name) Net IO" -Width 300 -Length 200 -Data1Label "Recv" -Data2Label "Trans" -Legend -Type "Line"
            
            WriteWordLine 0 0 ""           
        }
    }
    ElseIf($Text)
    {
        Line 0 "Host: $($VMHost.Name)"
        Line 0 ""
        Line 1 "Name:`t`t`t" $VMHost.Name
        Line 1 "ESXi Version:`t`t" $VMHost.Version
        Line 1 "ESXi Build:`t`t" $VMHost.Build
        Line 1 "Power State`t`t" $VMHost.PowerState
        Line 1 "Connection State:`t" $VMHost.ConnectionState
        Line 1 "Host Status:`t`t" $xStandAlone
        If (!$VMHost.IsStandAlone)
        {
            Line 1 "Parent Object:`t`t" $VMHost.Parent
        }
        If($VMHost.VMSwapfileDatastore)
        {
            Line 1 "VM Swapfile DS:`t" $VMHost.VMSwapfileDatastore.Name
        }
        Line 1 "Manufacturer:`t`t" $VMHost.Manufacturer
        Line 1 "Model:`t`t`t" $VMHost.Model
        Line 1 "CPU Type:`t`t" $VMHost.ProcessorType
        Line 1 "Maximum EVC Mode:`t" $VMHost.MaxEVCMode
        Line 1 "CPU Core Count:`t`t" $VMHost.NumCpu
        Line 1 "Hyperthreading:`t`t" $xHyperThreading
        Line 1 "CPU Power Policy:`t`t" $VMHost.ExtensionData.config.PowerSystemInfo.CurrentPolicy.ShortName
        Line 1 "Total Memory:`t`t$([decimal]::round($VMHost.MemoryTotalGB)) GB"
        Line 1 "SSH Policy:`t`t" (($xHostService) | Where {$_.Key -eq "TSM-SSH"}).Policy
        Line 1 "SSH Service Status:`t" $xSSHService
        Line 1 "Scratch Log location:`t" (($xHostAdvanced) | Where {$_.Name -eq "Syslog.global.logdir"}).Value
        Line 1 "Scratch Log server:`t" (($xHostAdvanced) | Where {$_.Name -eq "Syslog.global.loghost"}).Value
        Line 1 "NTP Service Policy:`t" (($xHostService) | Where {$_.Key -eq "ntpd"}).Policy
        Line 1 "NTP Service Status:`t" $xNTPService
        Line 1 "NFS Max Queue Depth:`t" (($xHostAdvanced) | Where {$_.Name -eq "NFS.MaxQueueDepth"}).Value
        Line 1 "NFS Max Volumes:`t" (($xHostAdvanced) | Where {$_.Name -eq "NFS.MaxQueueDepth"}).Value
        Line 1 "TCP IP Heap Size:`t" (($xHostAdvanced) | Where {$_.Name -eq "Net.TcpipHeapSize"}).Value
        Line 1 "TCP IP Heap Max:`t" (($xHostAdvanced) | Where {$_.Name -eq "Net.TcpipHeapMax"}).Value
        If ((($xHostService) | Where {$_.Key -eq "ntpd"}).Running)
        {
            Line 1 "NTP Servers" $xNTPServers
        }
        Line 0 ""
        If($xVMHostStorage)
        {
            Line 0 "Block Storage"
            Line 0 ""
            Line 1 "Model`tVendor`tCapacity`tRunTime Name`tMultipath`tIdentifier"
            ForEach($xLUN in ($xVMHostStorage.ScsiLun) | Where {$_.LunType -notlike "*cdrom*"})
            {
                Line 1 "$($xLUN.Model)`t$($xLUN.Vendor)`t$("{0:N2}" -f $xLUN.CapacityGB + " GB")`t$($xLUN.RuntimeName)`t$($xLUN.MultipathPolicy)`t$(truncate $xLUN.CanonicalName 28)"
            }
            Line 0 ""
        }
        If($VMHost.ConnectionState -like "*NotResponding*" -or $VMHost.PowerState -eq "PoweredOff")
        {
            Line 0 "Note: $($VMHost.Name) is not responding or is in an unknown state - data in this and other reports will be missing or inaccurate."
            Line 0 ""
        }
    }
    ElseIf($HTML)
    {
        $rowData = @()
        $colWidths = @("150px","200px")
        $rowData += @(,("Name",($htmlsilver -bor $htmlbold),$VMHost.Name,$htmlwhite))
        $rowData += @(,("ESXi Version",($htmlsilver -bor $htmlbold),$VMHost.Version,$htmlwhite))
        $rowData += @(,("ESXi Build",($htmlsilver -bor $htmlbold),$VMHost.Build,$htmlwhite))
        $rowData += @(,("Power State",($htmlsilver -bor $htmlbold),$VMHost.PowerState,$htmlwhite))
        $rowData += @(,("Connection State",($htmlsilver -bor $htmlbold),$VMHost.ConnectionState,$htmlwhite))
        $rowData += @(,("Host Status",($htmlsilver -bor $htmlbold),$xStandAlone,$htmlwhite))
        If (!$VMHost.IsStandAlone)
        {
            $rowData += @(,("Parent Object",($htmlsilver -bor $htmlbold),$VMHost.Parent,$htmlwhite))
        }
        If($VMHost.VMSwapfileDatastore)
        {
            $rowData += @(,("VM Swapfile DS",($htmlsilver -bor $htmlbold),$VMHost.VMSwapfileDatastore.Name,$htmlwhite))
        }
        $rowData += @(,("Manufacturer",($htmlsilver -bor $htmlbold),$VMHost.Manufacturer,$htmlwhite))
        $rowData += @(,("Model",($htmlsilver -bor $htmlbold),$VMHost.Model,$htmlwhite))
        $rowData += @(,("CPU Type",($htmlsilver -bor $htmlbold),$VMHost.ProcessorType,$htmlwhite))
        $rowData += @(,("Maximum EVC Mode",($htmlsilver -bor $htmlbold),$VMHost.MaxEVCMode,$htmlwhite))
        $rowData += @(,("CPU Core Count",($htmlsilver -bor $htmlbold),$VMHost.NumCpu,$htmlwhite))
        $rowData += @(,("Hyperthreading",($htmlsilver -bor $htmlbold),$xHyperThreading,$htmlwhite))
        $rowData += @(,("CPU Power Policy",($htmlsilver -bor $htmlbold),$VMHost.ExtensionData.config.PowerSystemInfo.CurrentPolicy.ShortName,$htmlwhite))
        $rowData += @(,("Total Memory",($htmlsilver -bor $htmlbold),"$([decimal]::round($VMHost.MemoryTotalGB)) GB",$htmlwhite))
        $rowData += @(,("SSH Policy",($htmlsilver -bor $htmlbold),(($xHostService) | Where {$_.Key -eq "TSM-SSH"}).Policy,$htmlwhite))
        $rowData += @(,("SSH Service Status",($htmlsilver -bor $htmlbold),$xSSHService,$htmlwhite))
        $rowData += @(,("Scratch Log location",($htmlsilver -bor $htmlbold),(($xHostAdvanced) | Where {$_.Name -eq "Syslog.global.logdir"}).Value,$htmlwhite))
        $rowData += @(,("Scratch Log Server",($htmlsilver -bor $htmlbold),(($xHostAdvanced) | Where {$_.Name -eq "Syslog.global.loghost"}).Value,$htmlwhite))
        $rowData += @(,("NTP Service Policy",($htmlsilver -bor $htmlbold),(($xHostService) | Where {$_.Key -eq "ntpd"}).Policy,$htmlwhite))
        $rowData += @(,("NTP Service Status",($htmlsilver -bor $htmlbold),$xNTPService,$htmlwhite))
        $rowData += @(,("NFS Max Queue Depth",($htmlsilver -bor $htmlbold),(($xHostAdvanced) | Where {$_.Name -eq "NFS.MaxQueueDepth"}).Value,$htmlwhite))
        $rowData += @(,("NFS Max Volumes",($htmlsilver -bor $htmlbold),(($xHostAdvanced) | Where {$_.Name -eq "NFS.MaxVolumes"}).Value,$htmlwhite))
        $rowData += @(,("TCP IP Heap Size",($htmlsilver -bor $htmlbold),(($xHostAdvanced) | Where {$_.Name -eq "Net.TcpipHeapSize"}).Value,$htmlwhite))
        $rowData += @(,("TCP IP Heap Max",($htmlsilver -bor $htmlbold),(($xHostAdvanced) | Where {$_.Name -eq "Net.TcpipHeapMax"}).Value,$htmlwhite))
        If ((($xHostService) | Where {$_.Key -eq "ntpd"}).Running)
        {
            $rowData += @(,("NTP Servers",($htmlsilver -bor $htmlbold),$xNTPServers,$htmlwhite))
        }

        FormatHTMLTable "Host: $($VMHost.Name)" -noHeadCols 2 -rowArray $rowData -fixedWidth $colWidths -tablewidth "350"
        WriteHTMLLine 0 1 ""
        If($xVMHostStorage)
        {
            $rowData = @()
            $columnHeaders = @("Model",($htmlsilver -bor $htmlbold),"Vendor",($htmlsilver -bor $htmlbold),"Capacity",($htmlsilver -bor $htmlbold),"RunTime Name",($htmlsilver -bor $htmlbold),"Multipath",($htmlsilver -bor $htmlbold),"Identifier",($htmlsilver -bor $htmlbold))
            ForEach($xLUN in ($xVMHostStorage.ScsiLun) | Where {$_.LunType -notlike "*cdrom*"})
            {
                $rowdata += @(,($xLun.Model,$htmlwhite,$xLun.Vendor,$htmlwhite,"$("{0:N2}" -f $xLUN.CapacityGB + " GB")",$htmlwhite,$xLUN.RuntimeName,$htmlwhite,$xLUN.MultipathPolicy,$htmlwhite,$xLUN.CanonicalName,$htmlwhite))
            }
            FormatHTMLTable "Block Storage" -rowArray $rowData -columnArray $columnHeaders
            WriteHTMLLine 0 0 " "
        }
        If($VMHost.ConnectionState -like "*NotResponding*" -or $VMHost.PowerState -eq "PoweredOff")
        {
            WriteHTMLLine 0 1 "Note: $($VMHost.Name) is not responding or is in an unknown state - data in this and other reports will be missing or inaccurate." "" $Null 0 $htmlitalics
            WriteHTMLLine 0 0 " "
        }
    }
}

Function ProcessClusters
{
    Write-Verbose "$(Get-Date): Processing VMware Clusters"
    If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "Clusters"
	}
	ElseIf($Text)
	{
		Line 0 "Clusters"
	}
    ElseIf($HTML)
    {
        WriteHTMLLine 1 0 "Clusters"
    }

    If($? -and ($Clusters))
    {
        ForEach($VMCluster in $Clusters)
        {
           OutputClusters $VMCluster
        }
    }
    ElseIf($? -and ($Clusters -eq $Null))
    {
        Write-Warning "There are no VMware Clusters"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There are no VMware Clusters"
		}
		ElseIf($Text)
		{
			Line 1 "There are no VMware Clusters"
		}
		ElseIf($HTML)
		{
           WriteHTMLLine 0 1 "There are no VMware Clusters"
		}
    }
    Else
    {
    	Write-Warning "Unable to retrieve VMware Clusters"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "Unable to retrieve VMware Clusters"
		}
		ElseIf($Text)
		{
			Line 1 "Unable to retrieve VMware Clusters"
		}
		ElseIf($HTML)
		{
            WriteHTMLLine 0 1 "Unable to retrieve VMware Clusters"
		}
    }


}

Function OutputClusters
{
    Param([object] $VMCluster)

    $xClusterHosts = (($VMHosts) | Where {$_.ParentId -eq $VMCLuster.Id} | Select -ExpandProperty Name) -join "`n"
    If($MSWord -or $PDF)
    {
        WriteWordLine 2 0 "Cluster: $($VMCluster.Name)"
        [System.Collections.Hashtable[]] $ScriptInformation = @()
        $ScriptInformation += @{ Data = "Name"; Value = $VMCluster.Name; }
        $ScriptInformation += @{ Data = "HA Enabled?"; Value = $VMCluster.HAEnabled; }
        If ($VMCluster.HAEnabled)
        {
            $ScriptInformation += @{ Data = "HA Admission Control Enabled?"; Value = $VMCluster.HAAdmissionControlEnabled; }
            $ScriptInformation += @{ Data = "HA Failover Level"; Value = $VMCluster.HAFailoverLevel; }
            $ScriptInformation += @{ Data = "HA Restart Priority"; Value = $VMCluster.HARestartPriority; }
            $ScriptInformation += @{ Data = "HA Isolation Response"; Value = $VMCluster.HAIsolationResponse; }
        }
        $ScriptInformation += @{ Data = "DRS Enabled?"; Value = $VMCluster.DrsEnabled; }
        If ($VMCluster.DrsEnabled)
        {
            $ScriptInformation += @{ Data = "DRS Automation Level"; Value = $VMCluster.DrsAutomationLevel; }
        }
        $ScriptInformation += @{ Data = "EVC Mode"; Value = $VMCluster.EVCMode; }
        If ($VMCluster.VsanEnabled)
        {
            $ScriptInformation += @{ Data = "VSAN Enabled?"; Value = $VMCluster.VsanEnabled; }
            $ScriptInformation += @{ Data = "VSAN Disk Claim Mode"; Value = $VMCluster.VsanDiskClaimMode; }
        }
        
        $Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 225;
		$Table.Columns.Item(2).Width = 200;

		# $Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""

        If ($VMCluster.DrsEnabled -and ($DRSRules | Where {$_.ClusterName -eq $VMCluster.Name}))
        {
            WriteWordLine 2 0 "DRS Rules and Groups"
            foreach($DRSRule in ($DRSRules | Where {$_.ClusterName -eq $VMCluster.Name}))
            {
                [System.Collections.Hashtable[]] $ScriptInformation = @()
                $ScriptInformation += @{ Data = "Rule Name"; Value = $DRSRule.RuleName; }
                $ScriptInformation += @{ Data = "Rule Type"; Value = $DRSRule.RuleType; }
                $ScriptInformation += @{ Data = "Rule Enabled"; Value = $DRSRule.bRuleEnabled; }
                If($DRSRule.bMandatory){$ScriptInformation += @{ Data = "Mandatory"; Value = $DRSRule.bMandatory; }}
                If($DRSRule.bKeepTogether){$ScriptInformation += @{ Data = "Keep Together"; Value = $DRSRule.bKeepTogether; }}
                If($DRSRule.VMNames){$ScriptInformation += @{ Data = "Virtual Machines"; Value = $DRSRule.VMNames; }}
                If($DRSRule.VMGroupName){$ScriptInformation += @{ Data = "VM Group"; Value = $DRSRule.VMGroupName; }}
                If($DRSRule.VMGroupMembers){$ScriptInformation += @{ Data = "Virtual Machines"; Value = $DRSRule.VMGroupMembers; }}
                If($DRSRule.AffineHostGrpName){$ScriptInformation += @{ Data = "Host Affinity Group"; Value = $DRSRule.AffineHostGrpName; }}
                If($DRSRule.AffineHostGrpMembers){$ScriptInformation += @{ Data = "Affinity Group Members"; Value = $DRSRule.AffineHostGrpMembers; }}
                If($DRSRule.AntiAffineHostGrpName){$ScriptInformation += @{ Data = "Host Anti Affinity Group"; Value = $DRSRule.AntiAffineHostGrpName; }}
                If($DRSRule.AntiAffineHostGrpMembers){$ScriptInformation += @{ Data = "Anti Affinity Group Members"; Value = $DRSRule.AntiAffineHostGrpMembers; }}

                $Table = AddWordTable -Hashtable $ScriptInformation `
                -Columns Data,Value -List -Format $wdTableGrid -AutoFit $wdAutoFitFixed

		        ## IB - Set the header row format
		        SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		        $Table.Columns.Item(1).Width = 225;
		        $Table.Columns.Item(2).Width = 200;

		        # $Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

		        FindWordDocumentEnd
		        $Table = $Null
		        WriteWordLine 0 0 ""

            }
        }

        If($xClusterHosts)
        {
            [System.Collections.Hashtable[]] $ScriptInformation = @()
            $ScriptInformation += @{ Data = "Hosts in $($VMCluster.Name)";}
            $ScriptInformation += @{ Data = $xClusterHosts;}

            $Table = AddWordTable -Hashtable $ScriptInformation `
		    -Columns Data `
		    -List `
		    -Format $wdTableGrid `
		    -AutoFit $wdAutoFitFixed;

		    ## IB - Set the header row format
		    SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		    $Table.Columns.Item(1).Width = 250;
		    # $Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

		    FindWordDocumentEnd
		    $Table = $Null
		    WriteWordLine 0 0 "" 
        }
        If($Chart)
        {
            $ClusterCpuAvg = get-stat -Entity $VMCluster.Name -Stat cpu.usagemhz.average -Start (Get-Date).AddDays(-7) -Finish (Get-Date) -IntervalMins 30
            AddStatsChart -StatData $ClusterCpuAvg -Type "Line" -Title "$($VMCluster.Name) CPU Percent" -Width 305 -Length 200

            $ClusterMemAvg = get-stat -Entity $VMCluster.Name -Stat mem.usage.average -Start (Get-Date).AddDays(-7) -Finish (Get-Date) -IntervalMins 30
            AddStatsChart -StatData $ClusterMemAvg -Type "Line" -Title "$($VMCluster.Name) Memory Percent" -Width 305 -Length 200

            WriteWordLine 0 0 ""
        }

    }
    ElseIf($HTML)
    {
        $rowdata = @()
        $colWidths = @("150px","200px")
        $rowdata += @(,("Name",($htmlsilver -bor $htmlbold),$VMCluster.Name,$htmlwhite))
        $rowdata += @(,("HA Enabled",($htmlsilver -bor $htmlbold),$VMCluster.HAEnabled,$htmlwhite))
        If ($VMCluster.HAEnabled)
        {
            $rowdata += @(,("HA Admission Control Enabled",($htmlsilver -bor $htmlbold),$VMCluster.HAAdmissionControlEnabled,$htmlwhite))
            $rowdata += @(,("HA Failover Level",($htmlsilver -bor $htmlbold),$VMCluster.HAFailoverLevel,$htmlwhite))
            $rowdata += @(,("HA Restart Priority",($htmlsilver -bor $htmlbold),$VMCluster.HARestartPriority,$htmlwhite))
            $rowdata += @(,("HA Isolation Response",($htmlsilver -bor $htmlbold),$VMCluster.HAIsolationResponse,$htmlwhite))
        }
        $rowdata += @(,("DRS Enabled",($htmlsilver -bor $htmlbold),$VMCluster.DrsEnabled,$htmlwhite))
        If ($VMCluster.DrsEnabled)
        {
            $rowdata += @(,("DRS Automation Level",($htmlsilver -bor $htmlbold),$VMCluster.DrsAutomationLevel,$htmlwhite))
        }
        $rowdata += @(,("EVC Mode",($htmlsilver -bor $htmlbold),$VMCluster.EVCMode,$htmlwhite))
        If ($VMCluster.VsanEnabled)
        {
            $rowdata += @(,("VSAN Enabled",($htmlsilver -bor $htmlbold),$VMCluster.VsanEnabled,$htmlwhite))
            $rowdata += @(,("VSAN Disk Claim Mode",($htmlsilver -bor $htmlbold),$VMCluster.VsanDiskClaimMode,$htmlwhite))
        }

        FormatHTMLTable "Cluster: $($VMCluster.Name)" -noHeadCols 2 -rowArray $rowdata -fixedWidth $colWidths -tablewidth "350"
        WriteHTMLLine 0 1 ""

        If ($VMCluster.DrsEnabled -and ($DRSRules | Where {$_.ClusterName -eq $VMCluster.Name}))
        {
            WriteHTMLLine 0 0 "DRS Rules and Groups" -options $htmlbold -fontSize 4
            foreach($DRSRule in ($DRSRules | Where {$_.ClusterName -eq $VMCluster.Name}))
            {
                $rowdata = @()
                $colWidths = @("150px","200px")
                $rowdata += @(,("Rule Name", ($htmlsilver -bor $htmlbold),$DRSRule.RuleName,$htmlwhite))
                $rowdata += @(,("Rule Type", ($htmlsilver -bor $htmlbold),$DRSRule.RuleType,$htmlwhite))
                $rowdata += @(,("Rule Enabled", ($htmlsilver -bor $htmlbold),$DRSRule.bRuleEnabled,$htmlwhite))
                If($DRSRule.bMandatory){$rowdata += @(,("Mandatory", ($htmlsilver -bor $htmlbold),$DRSRule.bMandatory,$htmlwhite))}
                If($DRSRule.bKeepTogether){$rowdata += @(,("Keep Together", ($htmlsilver -bor $htmlbold),$DRSRule.bKeepTogether,$htmlwhite))}
                If($DRSRule.VMNames){$rowdata += @(,("Virtual Machines", ($htmlsilver -bor $htmlbold),$DRSRule.VMNames,$htmlwhite))}
                If($DRSRule.VMGroupName){$rowdata += @(,("VM Group", ($htmlsilver -bor $htmlbold),$DRSRule.VMGroupName,$htmlwhite))}
                If($DRSRule.VMGroupMembers){$rowdata += @(,("Virtual Machines", ($htmlsilver -bor $htmlbold),$DRSRule.VMGroupMembers,$htmlwhite))}
                If($DRSRule.AffineHostGrpName){$rowdata += @(,("Host Affinity Group", ($htmlsilver -bor $htmlbold),$DRSRule.AffineHostGrpName,$htmlwhite))}
                If($DRSRule.AffineHostGrpMembers){$rowdata += @(,("Affinity Group Members", ($htmlsilver -bor $htmlbold),$DRSRule.AffineHostGrpMembers,$htmlwhite))}
                If($DRSRule.AntiAffineHostGrpName){$rowdata += @(,("Host Anti Affinity Group", ($htmlsilver -bor $htmlbold),$DRSRule.AntiAffineHostGrpName,$htmlwhite))}
                If($DRSRule.AntiAffineHostGrpMembers){$rowdata += @(,("Anti Affinity Group Members", ($htmlsilver -bor $htmlbold),$DRSRule.AntiAffineHostGrpMembers,$htmlwhite))}

                FormatHTMLTable "" -noHeadCols 2 -rowArray $rowdata -fixedWidth $colWidths -tablewidth "350"
                WriteHTMLLine 0 0 " "
            }
        }

        If ($xClusterHosts)
        {
            WriteHTMLLine 0 0 "Hosts in $($VMCluster.Name)" -options $htmlbold -fontSize 4
            ForEach ($cluHost in $xClusterHosts -split "`n"){WriteHTMLLine 0 1 $cluHost}
            WriteHTMLLine 0 1 ""
        }

    }
    ElseIf($Text)
    {
        Line 0 "Cluster: $($VMCluster.Name)"
        Line 0 ""
        Line 1 "HA Enabled:`t`t" $VMCluster.HAEnabled
        If ($VMCluster.HAEnabled)
        {
            Line 1 "HA Admission Control:`t" $VMCluster.HAAdmissionControlEnabled
            Line 1 "HA Failover Level:`t" $VMCluster.HAFailoverLevel
            Line 1 "HA Restart Priority:`t" $VMCluster.HARestartPriority
            Line 1 "HA Isolation Response:`t" $VMCluster.HAIsolationResponse
        }
        Line 1 "DRS Enabled:`t`t" $VMCluster.DrsEnabled
        If ($VMCluster.DrsEnabled)
        {
            Line 1 "DRS Automation Level:`t" $VMCluster.DrsAutomationLevel
        }
        Line 1 "EVC Mode:`t`t" $VMCluster.EVCMode
        If ($VMCluster.VsanEnabled)
        {
            Line 1 "VSAN Enabled:`t`t" $VMCluster.VsanEnabled
            Line 1 "VSAN Disk Claim Mode:`t" $VMCluster.VsanDiskClaimMode
        }
        Line 0 ""
        
        #Hosts in the cluster
        Line 0 "Hosts in $($VMCluster.Name)"
        Line 0 ""
        ForEach ($xHost in (($VMHosts) | Where {$_.ParentId -eq $VMCLuster.Id}).Name)
        {
            Line 1 $xHost
        }
        Line 0 ""
    }
}
#endregion

#region resource pools function
Function ProcessResourcePools
{
    Write-Verbose "$(Get-Date): Processing VMware Resource Pools"
    If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "Resource Pools"
	}
	ElseIf($Text)
	{
		Line 0 "Resource Pools"
	}
    ElseIf($HTML)
    {
        WriteHTMLLine 1 0 "Resource Pools"
    }

    If($? -and ($Resources))
    {
        ForEach($ResPool in $Resources)
        {
            OutputResourcePools $ResPool
        }
    }
    ElseIf($? -and ($Resources -eq $Null))
    {
        Write-Warning "There are no Resource Pools"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There are no Resource Pools"
		}
		ElseIf($Text)
		{
			Line 1 "There are no Resource Pools"
		}
		ElseIf($HTML)
		{
            WriteHTMLLine 0 1 "There are no Resource Pools"
		}
    }
    Else
    {
    	Write-Warning "Unable to retrieve Resource Pools"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "Unable to retrieve Resource Pools"
		}
		ElseIf($Text)
		{
			Line 1 "Unable to retrieve Resource Pools"
		}
		ElseIf($HTML)
		{
            WriteHTMLLine 0 1 "Unable to retrieve Resource Pools"
		}
    }

}
Function OutputResourcePools
{
    Param([object] $ResourcePool)
    If($Clusters.Name -contains $ResourcePool.Parent)
    {
        $xResourceParent = "$($ResourcePool.Parent) (Cluster Root)"
    }
    ElseIf($VMHosts.Name -contains $ResourcePool.Parent)
    {
        $xResourceParent = "$($ResourcePool.Parent) (Host Root)"
    }
    Else
    {
        $xResourceParent = "$($ResourcePool.Parent)"
    }
    If($ResourcePool.CpuLimitMHz -eq -1)
    {
        $xCpuLimit = "None"
    }
    Else
    {
        $xCpuLimit = "$($ResourcePool.CpuLimitMHz) MHz"
    }
    If($ResourcePool.MemReservationGB -eq -1)
    {
        $xMemRes = "None"
    }
    Else
    {
        If ($ResourcePool.MemReservationGB -lt 1)
        {
            $xMemRes = "$([decimal]::Round($ResourcePool.MemReservationMB)) MB"
        }
        $xMemRes = "$([decimal]::Round($ResourcePool.MemReservationGB)) GB"
    }      
    If($ResourcePool.MemLimitGB -eq -1)
    {
        $xMemLimit = "None"
    }
    Else
    {
        If ($ResourcePool.MemLimitGB -lt 1)
        {
            $xMemLimit = "$([decimal]::Round($ResourcePool.MemLimitMB)) MB"
        }
        $xMemLimit = "$([decimal]::Round($ResourcePool.MemLimitGB)) GB"
    }
    $xResPoolHosts = @(($VirtualMachines) | Where {$_.ResourcePoolId -eq $ResourcePool.Id} | Select Name | Sort Name)
    If($MSWord -or $PDF)
    {
        WriteWordLine 2 0 "Resource Pool: $($ResourcePool.Name)"
        [System.Collections.Hashtable[]] $ScriptInformation = @()
        $ScriptInformation += @{ Data = "Name"; Value = $ResourcePool.Name; }
        $ScriptInformation += @{ Data = "Parent Pool"; Value = $xResourceParent; }
        $ScriptInformation += @{ Data = "CPU Shares Level"; Value = $ResourcePool.CpuSharesLevel; }
        $ScriptInformation += @{ Data = "Number of CPU Shares"; Value = $ResourcePool.NumCpuShares; }
	    $ScriptInformation += @{ Data = "CPU Reservation"; Value = "$($ResourcePool.CpuReservationMHz) MHz"; }
	    $ScriptInformation += @{ Data = "CPU Limit"; Value = $xCpuLimit; }
        $ScriptInformation += @{ Data = "CPU Limit Expandable"; Value = $ResourcePool.CpuExpandableReservation; }
        $ScriptInformation += @{ Data = "Memory Shares Level"; Value = $ResourcePool.MemSharesLevel; }
        $ScriptInformation += @{ Data = "Number of Memory Shares"; Value = $ResourcePool.NumMemShares; }
	    $ScriptInformation += @{ Data = "Memory Reservation"; Value = $xMemRes; }
	    $ScriptInformation += @{ Data = "Memory Limit"; Value = $xMemLimit; }
        $ScriptInformation += @{ Data = "Memory Limit Expandable"; Value = $ResourcePool.MemExpandableReservation; }

        $Table = AddWordTable -Hashtable $ScriptInformation `
	    -Columns Data,Value `
	    -List `
	    -Format $wdTableGrid `
	    -AutoFit $wdAutoFitFixed;

	    ## IB - Set the header row format
	    SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

	    $Table.Columns.Item(1).Width = 205;
	    $Table.Columns.Item(2).Width = 200;

	    # $Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

	    FindWordDocumentEnd
	    $Table = $Null
	    WriteWordLine 0 0 ""
        
        If($xResPoolHosts)
        {
            If($xResPoolHosts.Count -gt 25)
            {
                WriteWordLine 0 0 "VMs in $($ResourcePool.Name)"
                BuildMultiColumnTable $xResPoolHosts.Name
                WriteWordLine 0 0 ""
            }
            Else
            {
                [System.Collections.Hashtable[]] $ScriptInformation = @()
                $ScriptInformation += @{ Data = "VMs in $($ResourcePool.Name)";}
                $ScriptInformation += @{ Data = ($xResPoolHosts.Name) -join "`n";}

                $Table = AddWordTable -Hashtable $ScriptInformation `
		        -Columns Data `
		        -List `
		        -Format $wdTableGrid `
		        -AutoFit $wdAutoFitFixed;

		        ## IB - Set the header row format
		        SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		        $Table.Columns.Item(1).Width = 280;

		        # $Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

		        FindWordDocumentEnd
		        $Table = $Null
		        WriteWordLine 0 0 ""  
            }
        }
    }
    ElseIf($HTML)
    {
        $rowdata = @()
        $colWidths = @("150px","200px")
        $rowdata += @(,("Parent Pool",($htmlsilver -bor $htmlbold),$xResourceParent,$htmlwhite))
        $rowdata += @(,("CPU Shares Level",($htmlsilver -bor $htmlbold),$ResourcePool.CpuSharesLevel,$htmlwhite))
        $rowdata += @(,("Number of CPU Shares",($htmlsilver -bor $htmlbold),$ResourcePool.NumCpuShares,$htmlwhite))
        $rowdata += @(,("CPU Reservation",($htmlsilver -bor $htmlbold),$xCpuLimit,$htmlwhite))
        $rowdata += @(,("CPU Limit",($htmlsilver -bor $htmlbold),$xCpuLimit,$htmlwhite))
        $rowdata += @(,("CPU Limit Expandable",($htmlsilver -bor $htmlbold),$ResourcePool.CpuExpandableReservation,$htmlwhite))
        $rowdata += @(,("Memory Shares Level",($htmlsilver -bor $htmlbold),$ResourcePool.MemSharesLevel,$htmlwhite))
        $rowdata += @(,("Number of Memory Shares",($htmlsilver -bor $htmlbold),$ResourcePool.NumMemShares,$htmlwhite))
        $rowdata += @(,("Memory Reservation",($htmlsilver -bor $htmlbold),$xMemRes,$htmlwhite))
        $rowdata += @(,("Memory Limit",($htmlsilver -bor $htmlbold),$xMemLimit,$htmlwhite))
        $rowdata += @(,("Memory Limit Expandable",($htmlsilver -bor $htmlbold),$ResourcePool.MemExpandableReservation,$htmlwhite))

        FormatHTMLTable "Resource Pool: $($ResourcePoolName)" -noHeadCols 2 -rowArray $rowdata -fixedWidth $colWidths -tablewidth "350"

        If($xResPoolHosts)
        {
            WriteHTMLLine 2 1 "VMs in $($ResourcePool.Name)"
            ForEach($xVM in (($VirtualMachines) | Where {$_.ResourcePoolId -eq $ResourcePool.Id} | Sort Name).Name)
            {
                WriteHTMLLine 0 2 $xVM
            }
            WriteHTMLLine 0 1 ""
        }
    }
    ElseIf($Text)
    {
        Line 0 "Resource Pool: $($ResourcePool.Name)"
        Line 0 ""
        Line 1 "Name:`t`t`t`t" $ResourcePool.Name
        Line 1 "Parent Pool:`t`t`t" $xResourceParent
        Line 1 "CPU Shares Level:`t`t" $ResourcePool.CpuSharesLevel
        Line 1 "Number of CPU Shares:`t`t" $ResourcePool.NumCpuShares
        Line 1 "CPU Reservation:`t`t$($ResourcePool.CpuReservationMHz) MHz"
        Line 1 "CPU Limit:`t`t`t" $xCpuLimit
        Line 1 "CPU Limit Expandable:`t`t" $ResourcePool.CpuExpandableReservation
        Line 1 "Memory Shares Level:`t`t" $ResourcePool.MemSharesLevel
        Line 1 "Number of Memory Shares:`t" $ResourcePool.NumMemShares
        Line 1 "Memory Reservation:`t`t" $xMemRes
        Line 1 "Memory Limit:`t`t`t" $xMemLimit
        Line 1 "Memory Limit Expandable:`t" $ResourcePool.MemExpandableReservation
        Line 0 ""

        If($xResPoolHosts)
        {
            Line 0 "VMs in $($ResourcePool.Name)"
            ForEach($xVM in (($VirtualMachines) | Where {$_.ResourcePoolId -eq $ResourcePool.Id} | Sort Name).Name)
            {
                Line 1 $xVM
            }
            Line 0 ""
        }
    }
}

#endregion

#region host networking and VMKernel ports functions
Function ProcessVMKPorts
{
    Write-Verbose "$(Get-Date): Processing VMkernel Ports"
    If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "VMKernel Ports"
	}
	ElseIf($Text)
	{
		Line 0 "VMKernel Ports"
	}
    ElseIf($HTML)
    {
        #WriteHTMLLine "VMKernel Ports"
    }
    $Script:VMKPortGroups = @()
    ForEach ($VMHost in ($VMHosts) | where {($_.ConnectionState -like "*Connected*") -or ($_.ConnectionState -like "*Maintenance*")})
    {
        If($MSWord -or $PDF){WriteWordLine 2 0 "VMKernel Ports on: $($VMHost.Name)"}
        If($Text){Line 0 ""; Line 0 "VMKernel Ports on: $($VMHost.Name)"}
        If($HTML){WriteHTMLLine 2 0 "VMKernel Ports on: $($VMHost.Name)"}
        $VMKPorts = $HostNetAdapters | Where {$_.DeviceName -Like "*vmk*" -and $_.VMHost -like $VMHost.Name} | Sort PortGroupName

        If($? -and ($VMKPorts))
        {
            ForEach ($VMK in $VMKPorts)
            {
                OutputVMKPorts $VMK
            }
        }
        ElseIf($? -and ($VMKPorts -eq $Null))
        {
            Write-Warning "There are no VMKernel ports"
		    If($MSWord -or $PDF)
		    {
			    WriteWordLine 0 1 "There are no VMKernel ports"
		    }
		    ElseIf($Text)
		    {
			    Line 1 "There are no VMKernel ports"
		    }
		    ElseIf($HTML)
		    {
                WriteHTMLLine 0 1 "There are no VMKernel ports"
		    }
        }
        Else
        {
    	    If(!($Export)){Write-Warning "Unable to retrieve VMKernel ports"}
		    If($MSWord -or $PDF)
		    {
			    WriteWordLine 0 1 "Unable to retrieve VMKernel ports"
		    }
		    ElseIf($Text)
		    {
			    Line 1 "Unable to retrieve VMKernel ports"
		    }
		    ElseIf($HTML)
		    {
                WriteHTMLLine 0 1 "Unable to retrieve VMKernel ports"
		    }
        }
    }
    
}

Function OutputVMKPorts
{
    Param([object] $VMK)

    $xSwitchDetail = $VirtualPortGroups | Where {$_.Name -like $VMK.PortGroupName} | Select -Unique
    $Script:VMKPortGroups += $VMK.PortGroupName

    If ($VMK.VMotionEnabled)
    {
        $xVMotionEnabled = "Yes"
    }
    Else
    {
        $xVMotionEnabled = "No"
    }
    If ($VMK.FaultToleranceLoggingEnabled)
    {
        $xFTLogging = "Yes"
    }
    Else
    {
        $xFTLogging = "No"
    }
    If ($VMK.ManagementTrafficEnabled)
    {
        $xMgmtTraffic = "Yes"
    }
    Else
    {
        $xMgmtTraffic = "No"
    }
    If ($VMK.DhcpEnabled)
    {
        $xIPAddressType = "DHCP"
    }
    Else
    {
        $xIPAddressType = "Static IP"
    }
    Switch ($xSwitchDetail.VLanId)
    {
        0
        {
            $xSwitchVLAN = "None"
        }
        4095
        {
            $xSwitchVLAN = "Trunk"
        }
        Default
        {
            $xSwitchVLAN = $xSwitchDetail.VLanId
        }
    }
    
    If($MSWord -or $PDF)
    {
        [System.Collections.Hashtable[]] $ScriptInformation = @()
        $ScriptInformation += @{ Data = "Port Name"; Value = $VMK.PortGroupName; }
        $ScriptInformation += @{ Data = "Port ID"; Value = $VMK.DeviceName; }
        $ScriptInformation += @{ Data = "MAC Address"; Value = $VMK.Mac; }
        $ScriptInformation += @{ Data = "IP Address Type"; Value = $xIPAddressType; }
        $ScriptInformation += @{ Data = "IP Address"; Value = $VMK.IP; }
        $ScriptInformation += @{ Data = "Subnet Mask"; Value = $VMK.SubnetMask; }
        $ScriptInformation += @{ Data = "VLAN ID"; Value = $xSwitchVLAN; }
        $ScriptInformation += @{ Data = "VMotion Traffic?"; Value = $xVMotionEnabled; }
        $ScriptInformation += @{ Data = "FT Logging Traffic?"; Value = $xFTLogging; }
        $ScriptInformation += @{ Data = "Management Traffic?"; Value = $xMgmtTraffic; }
        $ScriptInformation += @{ Data = "Parent vSwitch"; Value = $xSwitchDetail.VirtualSwitch; }

        $Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 190;
		$Table.Columns.Item(2).Width = 220;

		# $Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
    }
    ElseIf($HTML)
    {
        $rowdata = @()
        $colWidths = @("150px","200px")
        $rowdata += @(,("Port Name",($htmlsilver -bor $htmlbold),$VMK.PortGroupName,$htmlwhite))
        $rowdata += @(,("Port ID",($htmlsilver -bor $htmlbold),$VMK.DeviceName,$htmlwhite))
        $rowdata += @(,("MAC Address",($htmlsilver -bor $htmlbold),$VMK.Mac,$htmlwhite))
        $rowdata += @(,("IP Address Type",($htmlsilver -bor $htmlbold),$xIPAddressType,$htmlwhite))
        $rowdata += @(,("IP Address",($htmlsilver -bor $htmlbold),$VMK.IP,$htmlwhite))
        $rowdata += @(,("Subnet Mask",($htmlsilver -bor $htmlbold),$VMK.SubnetMask,$htmlwhite))
        $rowdata += @(,("VLAN ID",($htmlsilver -bor $htmlbold),$xSwitchVLAN,$htmlwhite))
        $rowdata += @(,("vMotion Traffic",($htmlsilver -bor $htmlbold),$xVMotionEnabled,$htmlwhite))
        $rowdata += @(,("FT Logging",($htmlsilver -bor $htmlbold),$xFTLogging,$htmlwhite))
        $rowdata += @(,("Management Traffic",($htmlsilver -bor $htmlbold),$xMgmtTraffic,$htmlwhite))
        $rowdata += @(,("Parent vSwitch",($htmlsilver -bor $htmlbold),$xSwitchDetail,$htmlwhite))
        
        FormatHTMLTable "" -noHeadCols 2 -rowArray $rowdata -fixedWidth $colWidths -tablewidth "350"
        WriteHTMLLine 0 0 " "
    }
    ElseIf($Text)
    {
        Line 1 "Port Name:`t`t" $VMK.PortGroupName
        Line 1 "Port ID:`t`t" $VMK.DeviceName
        Line 1 "MAC Address:`t`t" $VMK.Mac
        Line 1 "IP Address Type:`t" $xIPAddressType
        Line 1 "IP Address:`t`t" $VMK.IP
        Line 1 "Subnet Mask:`t`t" $VMK.SubnetMask
        Line 1 "VLAN ID:`t`t" $xSwichVLAN
        Line 1 "vMotion Traffic:`t" $xVMotionEnabled
        Line 1 "FT Logging Traffic:`t" $xFTLogging
        Line 1 "Management Traffic:`t" $xMgmtTraffic
        Line 1 "Parent vSwitch:`t`t" $xSwitchDetail.VirtualSwitch
        Line 0 ""
    }
}

Function ProcessHostNetworking
{
    Write-Verbose "$(Get-Date): Processing Host Networking"
    If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "Host Network Adapters"
	}
	ElseIf($Text)
	{
		Line 0 "Host Network Adapters"
	}
    ElseIf($HTML)
    {
        #WriteHTMLLine 1 0 "Host Network Adapters"
    }

    $NetArray = @()

    ForEach ($VMHost in ($VMHosts) | where {($_.ConnectionState -like "*Connected*") -or ($_.ConnectionState -like "*Maintenance*")})
    {
        If($Text){Line 0 ""; Line 0 "Host Network Adapters on: $($VMHost.Name)"}
        $HostNics = $HostNetAdapters | Where {$_.Name -notlike "*vmk*" -and $_.VMHost -like $VMHost.Name} | Sort Name
        ForEach($Nic in $HostNics)
        {
            $NetObject = New-Object psobject
            $NetObject | Add-Member -Name Hostname -MemberType NoteProperty -Value $VMHost.Name
            $NetObject | Add-Member -Name devName -MemberType NoteProperty -Value $Nic.DeviceName
            $NetObject | Add-Member -Name MAC -MemberType NoteProperty -Value $Nic.Mac
            $NetObject | Add-Member -Name Duplex -MemberType NoteProperty -Value $Nic.FullDuplex
            $NetObject | Add-Member -Name Speed -MemberType NoteProperty -Value $Nic.BitRatePerSec
            $netArray += $NetObject
        }

    }
    If($? -and ($netArray))
    {
        OutputHostNetworking $NetArray

    }
    ElseIf($? -and ($NetArray -eq $Null))
    {
        Write-Warning "There are no Host Network Adapters"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There are no Host Network Adapters"
		}
		ElseIf($Text)
		{
			    Line 1 "There are no Host Network Adapters"
		}
		ElseIf($HTML)
		{
            WriteHTMLLine 0 1 "There are no Host Network Adapters"
		}
    }
    Else
    {
    	    If(!($Export)){Write-Warning "Unable to retrieve Host Network Adapters"}
		    If($MSWord -or $PDF)
		    {
			    WriteWordLine 0 1 "Unable to retrieve Host Network Adapters"
		    }
		    ElseIf($Text)
		    {
			    Line 1 "Unable to retrieve Host Network Adapters"
		    }
		    ElseIf($HTML)
		    {
                WriteHTMLLine 0 1 "Unable to retrieve Host Network Adapters"
		    }
        }
}

Function OutputHostNetworking
{
    Param([object] $HostNic)

    If($MSWord -or $PDF)
    {
        ## Create an array of hashtables
	    [System.Collections.Hashtable[]] $HostNicWordTable = @();
	    ## Seed the row index from the second row
	    [int] $CurrentServiceIndex = 2;

        ForEach($xHostNic in $HostNic)
        {
            If ($xHostNic.Duplex)
            {
                $xDuplex = "Full Duplex"
            }
            Else
            {
                $xDuplex = "Half Duplex"
            }
            
            Switch($xHostNic.Speed)
            {
                0
                {
                    $xPortSpeed = "Down"
                    $xDuplex = ""
                }
                10
                {
                    $xPortSpeed = "10Mbps"
                }
                100
                {
                    $xPortSpeed = "100Mbps"
                }
                1000
                {
                    $xPortSpeed = "1Gbps"
                }
                10000
                {
                    $xPortSpeed = "10Gbps"
                }
                20000
                {
                    $xPortSpeed = "20Gbps"
                }
                40000
                {
                    $xPortSpeed = "40Gbps"
                }
                80000
                {
                    $xPortSpeed = "80Gbps"
                }
                100000
                {
                    $xPortSpeed = "100Gbps"
                }
            }

            ## Add the required key/values to the hashtable
	        $WordTableRowHash = @{
            Hostname = $xHostNic.Hostname;
            DeviceName = $xHostNic.devName;
            PortSpeed = $xPortSpeed;
            MACAddr = $xHostNic.Mac;
            Duplex = $xDuplex;
            }
            $HostNicWordTable += $WordTableRowHash;
            $CurrentServiceIndex++    
        }

        ## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
	    $Table = AddWordTable -Hashtable $HostNicWordTable `
	    -Columns HostName, DeviceName, PortSpeed, MACAddr, Duplex `
	    -Headers "Host", "Device Name", "Port Speed", "MAC Address", "Duplex" `
	    -Format $wdTableGrid `
	    -AutoFit $wdAutoFitContent;

	    ## IB - Set the header row format
	    SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

	    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

	    FindWordDocumentEnd
	    $Table = $Null

	    WriteWordLine 0 0 ""
    }
    ElseIf($Text)
    {
        ForEach($xHostNic in $HostNic)
        {
            If ($xHostNic.FullDuplex)
            {
                $xDuplex = "Full Duplex"
            }
            Else
            {
                $xDuplex = "Half Duplex"
            }
            
            Switch($xHostNic.BitRatePerSec)
            {
                0
                {
                    $xPortSpeed = "Down"
                    $xDuplex = ""
                }
                10
                {
                    $xPortSpeed = "10Mbps"
                }
                100
                {
                    $xPortSpeed = "100Mbps"
                }
                1000
                {
                    $xPortSpeed = "1Gbps"
                }
                10000
                {
                    $xPortSpeed = "10Gbps"
                }
                20000
                {
                    $xPortSpeed = "20Gbps"
                }
                40000
                {
                    $xPortSpeed = "40Gbps"
                }
                80000
                {
                    $xPortSpeed = "80Gbps"
                }
                100000
                {
                    $xPortSpeed = "100Gbps"
                }
            }

            Line 1 "Device Name:`t" $xHostNic.DeviceName
            Line 1 "Port Speed:`t" $xPortSpeed
            Line 1 "MAC Address:`t" $xHostNic.Mac
            If($xDuplex)
            {
                Line 1 "Duplex:`t" $xDuplex
            }
            Line 0 ""
        }
    }
    ElseIf($HTML)
    {
        $rowdata = @()
        ForEach($xHostNic in $HostNic)
        {
            If ($xHostNic.Duplex)
            {
                $xDuplex = "Full Duplex"
            }
            Else
            {
                $xDuplex = "Half Duplex"
            }
            
            Switch($xHostNic.Speed)
            {
                0
                {
                    $xPortSpeed = "Down"
                    $xDuplex = ""
                }
                10
                {
                    $xPortSpeed = "10Mbps"
                }
                100
                {
                    $xPortSpeed = "100Mbps"
                }
                1000
                {
                    $xPortSpeed = "1Gbps"
                }
                10000
                {
                    $xPortSpeed = "10Gbps"
                }
                20000
                {
                    $xPortSpeed = "20Gbps"
                }
                40000
                {
                    $xPortSpeed = "40Gbps"
                }
                80000
                {
                    $xPortSpeed = "80Gbps"
                }
                100000
                {
                    $xPortSpeed = "100Gbps"
                }
            }

            $columnHeaders = @("Host",($htmlSilver -bor $htmlbold),"Device Name",($htmlSilver -bor $htmlbold),"Port Speed",($htmlSilver -bor $htmlbold),"MAC Address",($htmlSilver -bor $htmlbold),"Duplex",($htmlSilver -bor $htmlbold))
            $rowdata += @(,($xHostNic.HostName,$htmlwhite,$xHostNic.devName,$htmlwhite,$xPortSpeed,$htmlwhite,$xHostNic.Mac,$htmlwhite,$xDuplex,$htmlwhite))
            
        }
        FormatHTMLTable "Host Network Adapters" -rowArray $rowdata -columnArray $columnHeaders
        WriteHTMLLine 0 0 " "
    }
}
#endregion

#region port groups and vswitch functions
Function ProcessVMPortGroups
{
    Write-Verbose "$(Get-Date): Processing VM Port Groups"
    If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "Virtual Machine Port Groups"
	}
	ElseIf($Text)
	{
        Line 0 ""
		Line 0 "Virtual Machine Port Groups"
	}
    ElseIf($HTML)
    {
        WriteHTMLLine 1 0 "Virtual Machine Port Groups"
    }

    If($? -and ($VirtualPortGroups))
    {
        ForEach($VMPortGroup in $VirtualPortGroups)
        {
            OutputVMPortGroups $VMPortGroup
        }
    }
    ElseIf($? -and ($VirtualPortGroups -eq $Null))
    {
        Write-Warning "There are no VM Port Groups"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There are no VM Port Groups"
		}
		ElseIf($Text)
		{
			Line 1 "There are no VM Port Groups"
		}
		ElseIf($HTML)
		{
            WriteHTMLLine 0 1 "There are no VM Port Groups"
		}
    }
    Else
    {
    	If(!($Export)){Write-Warning "Unable to retrieve VM Port Groups"}
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "Unable to retrieve VM Port Groups"
		}
		ElseIf($Text)
		{
			Line 1 "Unable to retrieve VM Port Groups"
		}
		ElseIf($HTML)
		{
            WriteHTMLLine 0 1 "Unable to retrieve VM Port Groups"
		}
    }

}

Function OutputVMPortGroups
{
    Param([object] $VMPortGroup)

    If ($Script:VMKPortGroups -notcontains $VMPortGroup.Name)
    {

        Switch ($VMPortGroup.VLanId)
        {
            0
            {
                $xPortVLAN = "None"
            }
            4095
            {
                $xPortVLAN = "Trunk"
            }
            Default
            {
                $xPortVLAN = $VMPortGroup.VLanId
            }
        }

        $xVMOnNetwork = @(($VMNetworkAdapters) | Where {$_.NetworkName -eq $VMPortGroup.Name} | Select Parent | Sort Parent | ForEach{$_.Parent})
            
        If($MSWord -or $PDF)
        {
            If($VMPortGroup.VLanId)
            {
                WriteWordLine 2 0 "VM Port Group: $($VMPortGroup.Name)"
                [System.Collections.Hashtable[]] $ScriptInformation = @()
                $ScriptInformation += @{ Data = "Parent vSwitch"; Value = $VMPortGroup.VirtualSwitch; }
                $ScriptInformation += @{ Data = "VLAN ID"; Value = $xPortVLAN; }

                $Table = AddWordTable -Hashtable $ScriptInformation `
		        -Columns Data,Value `
		        -List `
		        -Format $wdTableGrid `
		        -AutoFit $wdAutoFitFixed;

		        ## IB - Set the header row format
		        SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		        $Table.Columns.Item(1).Width = 225;
		        $Table.Columns.Item(2).Width = 200;

		        # $Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

		        FindWordDocumentEnd
		        $Table = $Null
		        WriteWordLine 0 0 ""

                If ($xVMOnNetwork)
                {
                    If($xVMOnNetwork.Count -gt 25)
                    {
                    WriteWordLine 0 0 "VMs in $($VMPortGroup.Name)"
                    BuildMultiColumnTable $xVMOnNetwork.Name
                    WriteWordLine 0 0 ""                    
                    }
                    Else
                    {
                    [System.Collections.Hashtable[]] $ScriptInformation = @()
                    $ScriptInformation += @{ Data = "VMs on $($VMPortGroup.Name)";}
                    $ScriptInformation += @{ Data = ($xVMOnNetwork.Name) -join "`n";}

                    $Table = AddWordTable -Hashtable $ScriptInformation `
		            -Columns Data `
		            -List `
		            -Format $wdTableGrid `
		            -AutoFit $wdAutoFitFixed;

		            ## IB - Set the header row format
		            SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		            $Table.Columns.Item(1).Width = 280;

		            # $Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

		            FindWordDocumentEnd
		            $Table = $Null
		            WriteWordLine 0 0 "" 
                    }
                }
            }

        }
        ElseIf($Text)
        {
            Line 1 "VM Port Group:`t`t$($VMPortGroup.Name)"
            Line 1 "Parent vSwitch:`t`t" $VMPortGroup.VirtualSwitch
            Line 1 "VLAN ID:`t`t" $xPortVLAN
            Line 0 ""

            If ($xVMOnNetwork)
            {
                Line 1 "VMs on $($VMPortGroup.Name)"
                ForEach($xVMNet in ($VMNetworkAdapters | Where {$_.NetworkName -eq $VMPortGroup.Name} | Select Parent | Sort Name))
                {
                    Line 1 $xVMNet.Parent
                }
            }
            Line 0 ""
        }
        ElseIf($HTML)
        {
            $rowData = @()
            $colWidths = @("150px","200px")
            $rowData += @(,("Parent vSwitch",($htmlsilver -bor $htmlbold),$VMPortGroup.Name,$htmlwhite))
            $rowData += @(,("VLAN ID",($htmlsilver -bor $htmlbold),$xPortVLAN,$htmlwhite))
            FormatHTMLTable "VM Port Group: $($VMPortGroup.Name)" -noHeadCols 2 -rowArray $rowData -fixedWidth $colWidths -tablewidth "350"
            WriteHTMLLine 0 0 " "

            If ($xVMOnNetwork)
            {
                WriteHTMLLine 2 1 "VMs on $($VMPortGroup.Name)"
                ForEach($xVMNet in ($VMNetworkAdapters | Where {$_.NetworkName -eq $VMPortGroup.Name} | Select Parent | Sort Name))
                {
                    WriteHTMLLine 0 1 $xVMNet.Parent
                }
                WriteHTMLLine 0 0 " "
            }
        }
    }

}

Function ProcessStandardVSwitch
{
    $DvSwitches = Get-VDSwitch 4>$Null
    If($DvSwitches)
    {
        ## DV Switches found - process them
        Write-Verbose "$(Get-Date): Processing DV Switching"
        If($MSWord -or $PDF)
        {
            $Selection.InsertNewPage()
            WriteWordLine 1 0 "DV Switching"
            OutputDVSwitching $DvSwitches
        }
        ElseIf($Text)
        {
            Line 0 "DV Switching"
        }
        ElseIf($HTML)
        {
            WriteHTMLLine 1 0 "DV Switching"
            OutputDVSwitching $DvSwitches
        }
    }
    Write-Verbose "$(Get-Date): Processing Standard vSwitching"
    If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "Standard vSwitching"
	}
	ElseIf($Text)
	{
		Line 0 "Standard vSwitching"
	}
    ElseIf($HTML)
    {
        
    }

    $vSwitchArray = @()
    ForEach ($VMHost in $VMHosts)
    {
        $stdVSwitchs = $VirtualSwitches | Where {$_.VMHost -like $VMHost.Name} | Sort Name
        ForEach ($vSwitch in $stdVSwitchs)
        {
            $switchObj = New-Object psobject
            $switchObj | Add-Member -Name HostName -MemberType NoteProperty -Value $VMHost.Name
            $switchObj | Add-Member -Name Name -MemberType NoteProperty -Value $vSwitch.Name
            $switchObj | Add-Member -Name NumPorts -MemberType NoteProperty -Value $vSwitch.NumPorts
            $switchObj | Add-Member -Name NumPortsAvailable -MemberType NoteProperty -Value $vSwitch.NumPortsAvailable
            $switchObj | Add-Member -Name Mtu -MemberType NoteProperty -Value $vSwitch.Mtu
            $switchObj | Add-Member -Name Nic -MemberType NoteProperty -Value $vSwitch.Nic
            $vSwitchArray += $switchObj

        }

    }

    If($? -and ($vSwitchArray))
    {
        OutputStandardVSwitch $vSwitchArray
    }
    ElseIf($? -and ($stdVSwitchs -eq $Null))
    {
            Write-Warning "There are no standard VSwitches configured on $($VMHost.Name)"
		    If($MSWord -or $PDF)
		    {
			    WriteWordLine 0 1 "There are no standard VSwitches configured on $($VMHost.Name)"
		    }
		    ElseIf($Text)
		    {
			    Line 1 "There are no standard VSwitches configured on $($VMHost.Name)"
		    }
		    ElseIf($HTML)
		    {
                WriteHTMLLine 0 1 "There are no standard VSwitches configured on $($VMHost.Name)"
		    }
        }
    Else
    {
    	    If(!($Export)){Write-Warning "Unable to retrieve standard VSwitches configured"}
		    If($MSWord -or $PDF)
		    {
			    WriteWordLine 0 1 "Unable to retrieve standard VSwitches configured"
		    }
		    ElseIf($Text)
		    {
			    Line 1 "Unable to retrieve standard VSwitches configured"
		    }
		    ElseIf($HTML)
		    {
                WriteHTMLLine 0 1 "Unable to retrieve standard VSwitches configured"
		    }
        }

}

Function OutputStandardVSwitch
{
    Param([object] $stdVSwitchs)

    If($MSWord -or $PDF)
    {
        ## Create an array of hashtables
	    [System.Collections.Hashtable[]] $switchWordTable = @();
	    ## Seed the row index from the second row
	    [int] $CurrentServiceIndex = 2;

        ForEach($stdVSwitch in $stdVSwitchs)
        {
            $xvSwitchNics = ($stdVSwitch.Nic) -join ", "

            ## Add the required key/values to the hashtable
	        $WordTableRowHash = @{
            VMHost = $stdVSwitch.HostName;
            switchName = $stdVSwitch.Name;
            NumPorts = $stdVSwitch.NumPorts;
            NumPortsAvail = $stdVSwitch.NumPortsAvailable;
            Mtu = $stdVSwitch.Mtu;
            Nics = $xvSwitchNics
            }

            $switchWordTable += $WordTableRowHash;
            $CurrentServiceIndex++;
        }

        $Table = AddWordTable -Hashtable $switchWordTable `
        -Columns VMHost, switchName, NumPorts, NumPortsAvail, Mtu, Nics `
        -Headers "Host", "vSwitch", "Total Ports", "Ports Available", "MTU", "Physical Adapters" `
        -Format $wdTableGrid `
        -AutoFit $wdAutoFitContent;

        SetWordTableAlternateRowColor $Table $wdColorGray05 "Second"
	    SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

	    # $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

	    FindWordDocumentEnd
	    $Table = $Null

	    WriteWordLine 0 0 ""
    }
    ElseIf($Text)
    {
        Line 1 "vSwitch Name:`t`t" $stdVSwitch.Name
        Line 1 "Total Ports:`t`t" $stdVSwitch.NumPorts
        Line 1 "Ports Available:`t" $stdVSwitch.NumPortsAvailable
        Line 1 "vSwitch MTU:`t`t" $stdVSwitch.Mtu
        Line 1 "Physical Host Adapters:`t" $xvSwitchNics
        Line 0 ""
    }
    ElseIf($HTML)
    {
        $rowdata = @()
        $columnHeaders = @("Host",($htmlsilver -bor $htmlbold),"vSwitch",($htmlsilver -bor $htmlbold),"Total Ports",($htmlsilver -bor $htmlbold),"Ports Available",($htmlsilver -bor $htmlbold),"MTU",($htmlsilver -bor $htmlbold),"Physical Adapters",($htmlsilver -bor $htmlbold))
        
        ForEach($stdVSwitch in $stdVSwitchs)
        {
            $xvSwitchNics = ($stdVSwitch.Nic) -join " "
            $rowdata += @(,($stdVSwitch.HostName,$htmlwhite,$stdVSwitch.Name,$htmlwhite,$stdVSwitch.NumPorts,$htmlwhite,$stdVSwitch.NumPortsAvailable,$htmlwhite,$stdVSwitch.Mtu,$htmlwhite,$xvSwitchNics,$htmlwhite))
        }
        FormatHTMLTable "Standard VSwitching" -rowArray $rowdata -columnArray $columnHeaders

    }

}

Function OutputDVSwitching
{
    Param([object] $dvSwitches)
    Write-Verbose "$(Get-Date): Gathering DV Switch data"

    $VdPortGroups = Get-VDPortgroup 4>$Null
    If($MSWord -or $PDF)
    {

        ## Create an array of hashtables
	    [System.Collections.Hashtable[]] $dvSwitchWordTable = @();
	    ## Seed the row index from the second row
	    [int] $CurrentServiceIndex = 2;

        ForEach($dvSwitch in $dvSwitches)
        {
            ## Add the required key/values to the hashtable
	        $WordTableRowHash = @{ 
            dvName = $dvSwitch.Name;
            dvVendor = $dvSwitch.Vendor;
            dvVersion = $dvSwitch.Version;
            dvUplink = $dvSwitch.NumUplinkPorts;
            dvMtu = $dvSwitch.Mtu
            }
	        ## Add the hash to the array
	        $dvSwitchWordTable += $WordTableRowHash;
	        $CurrentServiceIndex++;

        }

        ## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
	    $Table = AddWordTable -Hashtable $dvSwitchWordTable `
	    -Columns dvName, dvVendor, dvVersion, dvUplink, dvMtu `
	    -Headers "Switch Name", "Vendor", "Switch Version", "Uplink Ports", "Switch MTU" `
	    -Format $wdTableGrid `
	    -AutoFit $wdAutoFitContent;

	    ## IB - Set the header row format
	    SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

	    # $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

	    FindWordDocumentEnd
	    $Table = $Null

	    WriteWordLine 0 0 ""

        ## Create an array of hashtables
	    [System.Collections.Hashtable[]] $dvPortWordTable = @();
	    ## Seed the row index from the second row
	    [int] $CurrentServiceIndex = 2;

        Write-Verbose "$(Get-Date): Gathering DV Port data"
        ForEach($vdPortGroup in $VdPortGroups)
        {
            ForEach($vdPort in (Get-VDPort -VDPortgroup $VdPortGroup.Name 4>$null| Where {$_.ConnectedEntity -ne $null}))
            {
                #If($vdPort.ConnectedEntity -like "*vmk*"){$xPortName = "VMKernel"}Else{$xPortName = $vdPort.Name}
                If($VdPort.IsLinkUp){$xLinkUp = "Up"}Else{$xLinkUp = "Down"}

                ## Add the required key/values to the hashtable
	            $WordTableRowHash = @{ 
                hostName = $vdport.ProxyHost;
                entity = $vdport.ConnectedEntity;
                portGroup = $vdport.Portgroup;
                linkstatus = $xLinkUp;
                macAddr = $vdport.MacAddress;
                switch = $vdport.Switch
                }
	            ## Add the hash to the array
	            $dvPortWordTable += $WordTableRowHash;
	            $CurrentServiceIndex++;         
            
            }
        }

        ## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
	    $Table = AddWordTable -Hashtable $dvPortWordTable `
	    -Columns hostname, entity, portGroup, linkstatus, macAddr, switch `
	    -Headers "Host Name", "Entity", "Port Group", "Status", "MAC Address", "DV Switch" `
	    -Format $wdTableGrid `
	    -AutoFit $wdAutoFitContent;

	    ## IB - Set the header row format
	    SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

	    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

	    FindWordDocumentEnd
	    $Table = $Null

	    WriteWordLine 0 0 ""
    
    }

    If($HTML)
    {
        $rowData = @()
        $columnHeaders = @("Switch Name",($htmlsilver -bor $htmlbold),"Vendor",($htmlsilver -bor $htmlbold),"Switch Version",($htmlsilver -bor $htmlbold),"Uplink Ports",($htmlsilver -bor $htmlbold),"Switch MTU",($htmlsilver -bor $htmlbold))
        
        ForEach($dvSwitch in $dvSwitches)
        {
            $rowData += @(,($dvSwitch.Name,$htmlwhite,$dvSwitch.Vendor,$htmlwhite,$dvSwitch.Version,$htmlwhite,$dvSwitch.NumUplinkPorts,$htmlwhite,$dvSwitch.Mtu,$htmlwhite))
        }
        FormatHTMLTable "DV Switches" -rowArray $rowData -columnArray $columnHeaders
        WriteHTMLLine 0 0 " "

        $rowData = @()
        $columnHeaders = @("Host Name",($htmlsilver -bor $htmlbold),"Entity",($htmlsilver -bor $htmlbold),"Port Group",($htmlsilver -bor $htmlbold),"Status",($htmlsilver -bor $htmlbold),"MAC Address",($htmlsilver -bor $htmlbold),"DV Switch",($htmlsilver -bor $htmlbold))

        Write-Verbose "$(Get-Date): Gathering DV Port data"
        ForEach($vdPortGroup in $VdPortGroups)
        {
            ForEach($vdPort in (Get-VDPort -VDPortgroup $VdPortGroup.Name 4>$null| Where {$_.ConnectedEntity -ne $null}))
            {
                If($VdPort.IsLinkUp){$xLinkUp = "Up"}Else{$xLinkUp = "Down"}
                $rowData += @(,($vdport.ProxyHost,$htmlwhite,$vdport.ConnectedEntity,$htmlwhite,$vdport.Portgroup,$htmlwhite,$xLinkUp,$htmlwhite,$vdport.MacAddress,$htmlwhite,$vdport.Switch,$htmlwhite))
            }
        }
        FormatHTMLTable "DV SwitchPorts" -rowArray $rowData -columnArray $columnHeaders
        WriteHTMLLine 0 0 " "
    }

}
#endregion

#region storage and datastore functions

Function ProcessDatastores
{
    Write-Verbose "$(Get-Date): Processing Datastores"
    If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "VM Datastores"
	}
	ElseIf($Text)
	{
		Line 0 "VM Datastores"
	}
    ElseIf($HTML)
    {
        WriteHTMLLine 1 0 "Datastores"
    }
    ForEach ($Datastore in $Datastores)
    {
        OutputDatastores $Datastore
    }

}

Function OutputDatastores
{
    Param([object] $Datastore)

    If($Datastore.StorageIOControlEnabled)
    {
        $xSIOC = "Enabled"
    }
    Else
    {
        $xSIOC = "Disabled"
    }

    If($MSWord -or $PDF)
    {
        WriteWordLine 2 0 "Datastore: $($Datastore.Name)"
        [System.Collections.Hashtable[]] $ScriptInformation = @()
        $ScriptInformation += @{ Data = "Name"; Value = $Datastore.Name; }
        $ScriptInformation += @{ Data = "Type"; Value = $Datastore.Type; }
        $ScriptInformation += @{ Data = "Status"; Value = $Datastore.State; }
        $ScriptInformation += @{ Data = "Free Space"; Value = "$([decimal]::Round($Datastore.FreeSpaceGB)) GB"; }
        $ScriptInformation += @{ Data = "Capacity"; Value = "$([decimal]::Round($Datastore.CapacityGB)) GB"; }
        $ScriptInformation += @{ Data = "Storage IO Control"; Value = $xSIOC; }
        $ScriptInformation += @{ Data = "SIOC Threshold"; Value = "$($Datastore.CongestionThresholdMillisecond) ms"; }
        If($Datastore.Type -eq "NFS")
        {
            $ScriptInformation += @{ Data = "NFS Server"; Value = $Datastore.RemoteHost; }
            $ScriptInformation += @{ Data = "Share Path"; Value = $Datastore.RemotePath; }
        }
        If($Datastore.Type -eq "VMFS")
        {
            $ScriptInformation += @{ Data = "File System Version"; Value = $Datastore.FileSystemVersion; }
        }

        $Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 180;
		$Table.Columns.Item(2).Width = 260;

		# $Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
        
        #Hosts connected to this datastore
        $xHostsConnected = (($VMHosts) | Where {$_.DatastoreIdList -contains $Datastore.Id} | Select -ExpandProperty Name | Sort Name ) -join "`n"

        [System.Collections.Hashtable[]] $ScriptInformation = @()
        $ScriptInformation += @{ Data = "Hosts Connected to $($Datastore.Name)";}
        $ScriptInformation += @{ Data = $xHostsConnected;}

        $Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 280;

		# $Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 "" 
    }
    ElseIf($HTML)
    {
        $rowData = @()
        $colWidths = @("150px","200px")
        $rowData += @(,("Datastore",($htmlsilver -bor $htmlbold),$Datastore.Name,$htmlWhite))
        $rowData += @(,("Type",($htmlsilver -bor $htmlbold),$Datastore.Type,$htmlWhite))
        $rowData += @(,("Status",($htmlsilver -bor $htmlbold),$Datastore.State,$htmlWhite))
        $rowData += @(,("Free Space",($htmlsilver -bor $htmlbold),"$([decimal]::Round($Datastore.FreeSpaceGB)) GB",$htmlWhite))
        $rowData += @(,("Capacity",($htmlsilver -bor $htmlbold),"$([decimal]::Round($Datastore.CapacityGB)) GB",$htmlWhite))
        $rowData += @(,("Storage IO Control",($htmlsilver -bor $htmlbold),$xSIOC,$htmlWhite))
        $rowData += @(,("SIOC Threshold",($htmlsilver -bor $htmlbold),"$($Datastore.CongestionThresholdMillisecond) ms",$htmlWhite))
        If($Datastore.Type -eq "NFS")
        {
            $rowData += @(,("NFS Server",($htmlsilver -bor $htmlbold),$Datastore.RemoteHost,$htmlWhite))
            $rowData += @(,("Share Path",($htmlsilver -bor $htmlbold),$Datastore.RemotePath,$htmlWhite))
        }
        If($Datastore.Type -eq "VMFS")
        {
            $rowData += @(,("File System Version",($htmlsilver -bor $htmlbold),$Datastore.FileSystemVersion,$htmlWhite))
        }
        FormatHTMLTable $Datastore.Name -noHeadCols 2 -rowArray $rowData -fixedWidth $colWidths -tablewidth "350"
        WriteHTMLLine 0 1 ""
    }
    ElseIf($Text)
    {
        Line 0 "Datastore: $($Datastore.Name)"
        Line 0 ""
        Line 1 "Name:`t`t`t" $Datastore.Name
        Line 1 "Type:`t`t`t" $Datastore.Type
        Line 1 "Status:`t`t`t" $Datastore.State
        Line 1 "Free Space:`t`t$([decimal]::Round($Datastore.FreeSpaceGB)) GB"
        Line 1 "Capacity:`t`t$([decimal]::Round($Datastore.CapacityGB)) GB"
        Line 1 "Storage IO Control:`t" $xSIOC
        Line 1 "SIOC Threshold:`t`t$($Datastore.CongestionThresholdMillisecond) ms"
        If($Datastore.Type -eq "NFS")
        {
            Line 1 "NFS Server:`t`t" $Datastore.RemoteHost
            Line 1 "Share Path:`t`t" $Datastore.RemotePath
        }
        If($Datastore.Type -eq "VMFS")
        {
            Line 1 "File System Version:`t" $Datastore.FileSystemVersion
        }
        Line 0 ""
    }
}

#endregion

#region virtual machine functions

Function ProcessVirtualMachines
{
    Write-Verbose "$(Get-Date): Processing Virtual Machines"
    If($MSWord -or $PDF)
	{
		$Selection.InsertNewPage()
		WriteWordLine 1 0 "Virtual Machines"
        $First = $True
	}
	ElseIf($Text)
	{
		Line 0 "Virtual Machines"
	}


    If($? -and ($VirtualMachines))
    {
        $First = $True
        ForEach($VM in $VirtualMachines)
        {
            If(!$First -and !$Export){$Selection.InsertNewPage()}
            OutputVirtualMachines $VM
            $First = $False
        }
    }
    ElseIf($? -and ($VirtualMachines -eq $Null))
    {
        Write-Warning "There are no Virtual Machines"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There are no Virtual Machines"
		}
		ElseIf($Text)
		{
			Line 1 "There are no Virtual Machines"
		}
		ElseIf($HTML)
		{
		}
    }
    Else
    {
    	Write-Warning "Unable to retrieve Virtual Machines"
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "Unable to retrieve Virtual Machines"
		}
		ElseIf($Text)
		{
			Line 1 "Unable to retrieve Virtual Machines"
		}
		ElseIf($HTML)
		{
		}
    }

}

Function OutputVirtualMachines
{
    Param([object] $VM)

    If($VM.MemoryGB -lt 1)
    {
        $xMemAlloc = "$($VM.MemoryMB) MB"
    }
    Else
    {
        $xMemAlloc = "$($VM.MemoryGB) GB"
    }
    If($VM.Guest.OSFullName)
    {
        $xGuestOS = $VM.Guest.OSFullName
    }
    Else
    {
        $xGuestOS = $VM.GuestId
    }
    If($VM.VApp)
    {
        $xParentFolder = "VAPP: $($VM.VApp)"
        $xParentResPool = "VAPP: $($VM.VApp)"
    }
    Else
    {
        $xParentFolder = $VM.Folder
        $xParentResPool = $VM.ResourcePool
    }
    If($VM.ExtensionData.Config.CpuAllocation.Limit -eq -1)
    {$xCpuLimit = 'Unlimited'}
    Else
    {$xCpuLimit = "$($VM.ExtensionData.Config.CpuAllocation.Limit) MHz"}
    If($VM.ExtensionData.Config.MemoryAllocation.Limit -eq -1)
    {$xMemLimit = 'Unlimited'}
    Else
    {$xMemLimit = "$($VM.ExtensionData.Config.MemoryAllocation.Limit) MB"}

    If($VM.Guest.State -eq "Running")
    {
        $xVMDetail = $True
        If($Export){$VM.Guest | Export-Clixml .\Export\VMDetail\$($VM.Name)-Detail.xml 4>$Null}
        If($Import){$GuestImport = Import-Clixml .\Export\VMDetail\$($VM.Name)-Detail.xml}
    }

    If($MSWord -or $PDF)
    {
        WriteWordLine 2 0 "VM: $($VM.Name)"
        [System.Collections.Hashtable[]] $ScriptInformation = @()
        $ScriptInformation += @{ Data = "Name"; Value = $VM.Name; }
        $ScriptInformation += @{ Data = "Guest OS"; Value = $xGuestOS; }
        $ScriptInformation += @{ Data = "VM Hardware Version"; Value = $VM.Version; }
        $ScriptInformation += @{ Data = "Power State"; Value = $VM.PowerState; }
        $ScriptInformation += @{ Data = "Guest Tools Status"; Value = $VM.Guest.State; }
        If($VM.Description)
        {
            $ScriptInformation += @{ Data = "Description"; Value = $VM.Description.Replace("`n"," "); }
        }
        If($VM.Notes)
        {
            $ScriptInformation += @{ Data = "Notes"; Value = $VM.Notes.Replace("`n"," "); }
        }
        $ScriptInformation += @{ Data = "Guest Tools Time Sync"; Value = $VM.ExtensionData.Config.Tools.SyncTimeWithHost; }
        $ScriptInformation += @{ Data = "Current Host"; Value = $VM.Host; }
        $ScriptInformation += @{ Data = "Parent Folder"; Value = $xParentFolder; }
        $ScriptInformation += @{ Data = "Parent Resource Pool"; Value = $xParentResPool; }
        If($VM.VApp)
        {
            $ScriptInformation += @{ Data = "Part of a VApp"; Value = $VM.VApp; }
        }
        $ScriptInformation += @{ Data = "vCPU Sockets"; Value = $VM.NumCPU/$VM.ExtensionData.Config.Hardware.NumCoresPerSocket; }
        $ScriptInformation += @{ Data = "vCPU Cores per Socket"; Value = $VM.ExtensionData.COnfig.Hardware.NumCoresPerSocket; }
        $ScriptInformation += @{ Data = "vCPU Total"; Value = $VM.NumCpu; }
        $ScriptInformation += @{ Data = "CPU Resources"; Value = "$($VM.VMResourceConfiguration.CpuSharesLevel) - $($VM.VMResourceConfiguration.NumCpuShares)"; }
        $ScriptInformation += @{ Data = "CPU Reservation"; Value = "$($VM.ExtensionData.Config.CpuAllocation.Reservation) Mhz"; }
        $ScriptInformation += @{ Data = "CPU Resource Limit"; Value = $xCpuLimit; }
        $ScriptInformation += @{ Data = "RAM Allocation"; Value = $xMemAlloc; }
        $ScriptInformation += @{ Data = "RAM Resources"; Value = "$($VM.VMResourceConfiguration.MemSharesLevel) - $($VM.VMResourceConfiguration.NumMemShares)"; }
        $ScriptInformation += @{ Data = "RAM Reservation"; Value = "$($VM.ExtensionData.Config.MemoryAllocation.Reservation) MB"; }
        $ScriptInformation += @{ Data = "RAM Resource Limit"; Value = $xMemLimit; }
        $xNicCount = 0
        ForEach($VMNic in $VM.NetworkAdapters)
        {
            $xNicCount += 1
            $ScriptInformation += @{ Data = "Network Adapter $($xNicCount)"; Value = $VMNic.Type; }
            $ScriptInformation += @{ Data = "     Port Group"; Value = $VMNic.NetworkName; }
            $ScriptInformation += @{ Data = "     MAC Address"; Value = $VMNic.MacAddress; }
            If($Import){$xVMGuestNics = $GuestImport.Nics}Else{$xVMGuestNics = $VM.Guest.Nics}
            If($xVMDetail){$ScriptInformation += @{ Data = "     IP Address"; Value = (($xVMGuestNics | Where {$_.Device -like "Network Adapter $($xNicCount)"}).IPAddress |Where {$_ -notlike "*:*"}) -join ", "; }}
        }
        $ScriptInformation += @{ Data = "Storage Allocation"; Value = "$([decimal]::Round($VM.ProvisionedSpaceGB)) GB"; }
        $ScriptInformation += @{ Data = "Storage Usage"; Value = "{0:N2}" -f $VM.UsedSpaceGB + " GB"; }
        If($xVMDetail)
        {
            If($Import){$xVMDisks = $GuestImport.Disks}Else{$xVMDisks = $VM.Guest.Disks}
            ForEach($VMVolume in ($xVMDisks | sort Path))
            {
                $ScriptInformation += @{ Data = "Guest Volume Path"; Value = $VMVolume.Path; }
                $ScriptInformation += @{ Data = "     Capacity"; Value = "{0:N2}" -f $VMVolume.CapacityGB + " GB"; }
                $ScriptInformation += @{ Data = "     Free Space"; Value = "{0:N2}" -f $VMVolume.FreeSpaceGB + " GB"; }
            }
        }
        $xDiskCount = 0
        ForEach($VMDisk in $VM.HardDisks)
        {
            $xDiskCount += 1
            $ScriptInformation += @{ Data = "Hard Disk $($xDiskCount)"; Value = "{0:N2}" -f $VMDisk.CapacityGB + " GB"; }
            $ScriptInformation += @{ Data = "     Datastore"; Value = $VMDisk.Filename.Substring($VMDisk.Filename.IndexOf("[")+1,$VMDisk.Filename.IndexOf("]")-1); }
            $ScriptInformation += @{ Data = "     Disk Path"; Value = $VMDisk.Filename.Substring($VMDisk.Filename.IndexOf("]")+2); }
            $ScriptInformation += @{ Data = "     Format"; Value = $VMDisk.StorageFormat; }
            $ScriptInformation += @{ Data = "     Type"; Value = $VMDisk.DiskType; }
            $ScriptInformation += @{ Data = "     Persistence"; Value = $VMDisk.Persistence; }
        }
        If(($Snapshots) | Where {$_.VM -like $VM.Name}){$ScriptInformation += @{ Data = "VM has Snapshots"; Value = (($Snapshots) | Where {$_.VM -like $VM.Name}).Count; }}

        $Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 180;
		$Table.Columns.Item(2).Width = 260;

		# $Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
        
        If($VM.Guest.State -eq "Running" -and $Chart)
        {
            $VMCpuAvg = get-stat -Entity $VM.Name -Stat cpu.usage.average -Start (Get-Date).AddDays(-7) -Finish (Get-Date) -IntervalMins 30
            AddStatsChart -StatData $VMCpuAvg -Type "Line" -Title "$($VM.Name) CPU Percent" -Width 250 -Length 200

            $VMMemAvg = get-stat -Entity $VM.Name -Stat mem.usage.average -Start (Get-Date).AddDays(-7) -Finish (Get-Date) -IntervalMins 30
            AddStatsChart -StatData $VMMemAvg -Type "Line" -Title "$($VM.Name) Memory Percent" -Width 250 -Length 200

            $VMdiskWrite = get-stat -Entity $VM.Name -Stat "virtualDisk.write.average" -Start (Get-Date).AddDays(-7) -Finish (Get-Date) -IntervalMins 30 | Where {$_.Instance -like ""}
            $VMdiskread = get-stat -Entity $VM.Name -Stat "virtualDisk.read.average" -Start (Get-Date).AddDays(-7) -Finish (Get-Date) -IntervalMins 30 | Where {$_.Instance -like ""}
            AddStatsChart -StatData $VMdiskWrite -StatData2 $VMdiskread -Title "$($VM.Name) Disk IO" -Width 300 -Length 200 -Data1Label "Write IO" -Data2Label "Read IO" -Legend -Type "Line"

            $VMNetRec = get-stat -Entity $VM.Name -Stat "net.received.average" -Start (Get-Date).AddDays(-7) -Finish (Get-Date) -IntervalMins 30 | Where {$_.Instance -like ""}
            $VMNetTrans = get-stat -Entity $VM.Name -Stat "net.transmitted.average" -Start (Get-Date).AddDays(-7) -Finish (Get-Date) -IntervalMins 30 | Where {$_.Instance -like ""}
            AddStatsChart -StatData $VMNetRec -StatData2 $VMNetTrans -Title "$($VM.Name) Net IO" -Width 300 -Length 200 -Data1Label "Recv" -Data2Label "Trans" -Legend -Type "Line"
        } 
    
    }
    ElseIf($Text)
    {
        Line 0 "VM: $($VM.Name)"
        Line 1 "Name:`t`t`t" $VM.Name
        Line 1 "Guest OS:`t`t" $xGuestOS
        Line 1 "VM Hardware Version:`t" $VM.Version
        Line 1 "Power State:`t`t" $VM.PowerState
        Line 1 "Guest Tools Status:`t" $VM.Guest.State
        If($VM.Description){Line 1 "Description:`t`t" $VM.Description.Replace("`n"," ")}
        If($VM.Notes){Line 1 "Notes:`t`t`t" $VM.Notes.Replace("`n"," ")}
        Line 1 "Guest Tools Time Sync:`t" $VM.ExtensionData.Config.Tools.SyncTimeWithHost
        Line 1 "Current Host:`t`t" $VM.Host
        Line 1 "Parent Folder:`t`t" $xParentFolder
        Line 1 "Parent Resource Pool:`t" $xParentResPool
        If($VM.VApp){Line 1 "Part of a VApp:`t" $VM.VApp}
        Line 1 "vCPU Sockets:`t`t" ($VM.NumCPU/$VM.ExtensionData.Config.Hardware.NumCoresPerSocket)
        Line 1 "vCPU Cores per Socket:" $VM.ExtensionData.Config.Hardware.NumCoresPerSocket
        Line 1 "vCPU Total:`t`t" $VM.NumCpu
        Line 1 "CPU Resources:`t`t $($VM.VMResourceConfiguration.CpuSharesLevel) - $($VM.VMResourceConfiguration.NumCpuShares)"
        Line 1 "CPU Reservation:`t`t $($VM.ExtensionData.Config.CpuAllocation.Reservation) Mhz"
        Line 1 "CPU Resource Limit`t:" $xCpuLimit
        Line 1 "RAM Allocation:`t`t" $xMemAlloc
        Line 1 "RAM Resources:`t`t $($VM.VMResourceConfiguration.MemSharesLevel) - $($VM.VMResourceConfiguration.NumMemShares)"
        Line 1 "RAM Reservation`t`t $($VM.ExtensionData.Config.MemoryAllocation.Reservation) MB"
        Line 1 "RAM Resource Limit`t" $xMemLimit
        $xNicCount = 0
        ForEach($VMNic in $VM.NetworkAdapters)
        {
            $xNicCount += 1
            Line 1 "Network Adapter $($xNicCount):`t" $VMNic.Type
            Line 1 "Port Group:`t`t" $VMNic.NetworkName
            Line 1 "MAC Address:`t`t" $VMNic.MacAddress
            If($xVMDetail){Line 1 "IP Address:`t`t" (($VM.Guest.Nics | Where {$_.Device -like "Network Adapter $($xNicCount)"}).IPAddress |Where {$_ -notlike "*:*"}) -join ", "}
        }
        Line 1 "Storage Allocation:`t$([decimal]::Round($VM.ProvisionedSpaceGB)) GB"
        Line 1 "Storage Usage:`t`t" $("{0:N2}" -f $VM.UsedSpaceGB + " GB")
        If($xVMDetail)
        {
            ForEach($VMVolume in $VM.Guest.Disks)
            {
                Line 1 "Guest Volume Path:`t" $VMVolume.Path
                Line 1 "Capacity:`t`t" $("{0:N2}" -f $VMVolume.CapacityGB + " GB")
                Line 1 "Free Space:`t`t" $("{0:N2}" -f $VMVolume.FreeSpaceGB + " GB")
            }
        }
        $xDiskCount = 0
        ForEach($VMDisk in $VM.HardDisks)
        {
            $xDiskCount += 1
            Line 1 "Hard Disk $($xDiskCount):`t`t" $("{0:N2}" -f $VMDisk.CapacityGB + " GB")
            Line 1 "Datastore:`t`t" $VMDisk.Filename.Substring($VMDisk.Filename.IndexOf("[")+1,$VMDisk.Filename.IndexOf("]")-1)
            Line 1 "Disk Path:`t`t" $VMDisk.Filename.Substring($VMDisk.Filename.IndexOf("]")+2)
            Line 1 "Format:`t`t`t" $VMDisk.StorageFormat
            Line 1 "Type:`t`t`t" $VMDisk.DiskType
            Line 1 "Persistence:`t`t" $VMDisk.Persistence
        }
        If(($Snapshots) | Where {$_.VM -like $VM.Name}){Line 1 "VM has Snapshots:`t" (($Snapshots) | Where {$_.VM -like $VM.Name}).Count}
        Line 0 ""
    }
    ElseIf($HTML)
    {
        $rowdata = @()
        $colWidths = @("150px","200px")
        $rowdata += @(,("Name",($htmlsilver -bor $htmlbold),$VM.Name,$htmlwhite))
        $rowdata += @(,("Guest OS",($htmlsilver -bor $htmlbold),$xGuestOS,$htmlwhite))
        $rowdata += @(,("VM Hardware Version",($htmlsilver -bor $htmlbold),$VM.Version,$htmlwhite))
        $rowdata += @(,("Power State",($htmlsilver -bor $htmlbold),$VM.PowerState,$htmlwhite))
        $rowdata += @(,("Guest Tools Status",($htmlsilver -bor $htmlbold),$VM.Guest.State,$htmlwhite))
        If($VM.Description)
        {
            $rowdata += @(,("Description",($htmlsilver -bor $htmlbold),$VM.Description.Replace("`n"," "),$htmlwhite))
        }
        If($VM.Notes)
        {
            $rowdata += @(,("Notes",($htmlsilver -bor $htmlbold),$VM.Notes.Replace("`n"," "),$htmlwhite))
        }
        $rowdata += @(,("Guest Tools Time Sync",($htmlsilver -bor $htmlbold),$VM.ExtensionData.Config.Tools.SyncTimeWithHost,$htmlwhite))
        $rowdata += @(,("Current Host",($htmlsilver -bor $htmlbold),$VM.Host,$htmlwhite))
        $rowdata += @(,("Parent Folder",($htmlsilver -bor $htmlbold),$xParentFolder,$htmlwhite))
        $rowdata += @(,("Parent Resource Pool",($htmlsilver -bor $htmlbold),$xParentResPool,$htmlwhite))
        If($VM.VApp)
        {
            $rowdata += @(,("Part of a VApp",($htmlsilver -bor $htmlbold),$VM.VApp,$htmlwhite))
        }
        $rowdata += @(,("vCPU Sockets",($htmlsilver -bor $htmlbold),($VM.NumCPU/$VM.ExtensionData.Config.Hardware.NumCoresPerSocket),$htmlwhite))
        $rowdata += @(,("vCPU Cores per Socket",($htmlsilver -bor $htmlbold),$VM.ExtensionData.Config.Hardware.NumCoresPerSocket,$htmlwhite))
        $rowdata += @(,("vCPU Total",($htmlsilver -bor $htmlbold),$VM.NumCpu,$htmlwhite))
        $rowdata += @(,("CPU Resources",($htmlsilver -bor $htmlbold),"$($VM.VMResourceConfiguration.CpuSharesLevel) - $($VM.VMResourceConfiguration.NumCpuShares)",$htmlwhite))
        $rowdata += @(,("CPU Reservation",($htmlsilver -bor $htmlbold),"$($VM.ExtensionData.Config.CpuAllocation.Reservation) Mhz",$htmlwhite))
        $rowdata += @(,("CPU Resource Limit",($htmlsilver -bor $htmlbold),$xCpuLimit,$htmlwhite))
        $rowdata += @(,("RAM Allocation",($htmlsilver -bor $htmlbold),$xMemAlloc,$htmlwhite))
        $rowdata += @(,("RAM Resources",($htmlsilver -bor $htmlbold),"$($VM.VMResourceConfiguration.MemSharesLevel) - $($VM.VMResourceConfiguration.NumMemShares)",$htmlwhite))
        $rowdata += @(,("RAM Reservation",($htmlsilver -bor $htmlbold),"$($VM.ExtensionData.Config.MemoryAllocation.Reservation) MB",$htmlwhite))
        $rowdata += @(,("RAM Resource Limit",($htmlsilver -bor $htmlbold),$xMemLimit,$htmlwhite))
        $xNicCount = 0
        Foreach($VMNic in $VM.NetworkAdapters)
        {
            $xNicCount += 1
            $rowdata += @(,("Network Adapter $($xNicCount)",($htmlsilver -bor $htmlbold),$VMNic.Type,$htmlwhite))
            $rowdata += @(,("Port Group",($htmlsilver -bor $htmlitalics),$VMNic.NetworkName,$htmlwhite))
            $rowdata += @(,("MAC Address",($htmlsilver -bor $htmlitalics),$VMNic.MacAddress,$htmlwhite))
            If($Import){$xVMGuestNics = $GuestImport.Nics}Else{$xVMGuestNics = $VM.Guest.Nics}
            If($xVMDetail){$rowdata += @(,("IP Address",($htmlsilver -bor $htmlitalics),$((($xVMGuestNics | Where {$_.Device -like "Network Adapter $($xNicCount)"}).IPAddress |Where {$_ -notlike "*:*"}) -join ", "),$htmlwhite))
        }

        }
        $rowdata += @(,("Storage Allocation",($htmlsilver -bor $htmlbold),"$([decimal]::Round($VM.ProvisionedSpaceGB)) GB",$htmlwhite))
        $rowdata += @(,("Storage Usage",($htmlsilver -bor $htmlbold),$("{0:N2}" -f $VM.UsedSpaceGB + " GB"),$htmlwhite))
        If($xVMDetail)
        {
            If($Import){$xVMDisks = $GuestImport.Disks}Else{$xVMDisks = $VM.Guest.Disks}
            foreach($VMVolume in $xVMDisks)
            {
                $rowdata += @(,("Guest Volume Path",($htmlsilver -bor $htmlbold),$VMVolume.Path,$htmlwhite))
                $rowdata += @(,("Capacity",($htmlsilver -bor $htmlitalics),$("{0:N2}" -f $VMVolume.CapacityGB + " GB"),$htmlwhite))
                $rowdata += @(,("Free Space",($htmlsilver -bor $htmlitalics),$("{0:N2}" -f $VMVolume.FreeSpaceGB + " GB"),$htmlwhite))
            }
        }
        $xDiskCount = 0
        foreach($VMDisk in $VM.HardDisks)
        {
            $xDiskCount += 1
            $rowdata += @(,("Hard Disk $($xDiskCount)",($htmlsilver -bor $htmlbold),$("{0:N2}" -f $VMDisk.CapacityGB + " GB"),$htmlwhite))
            $rowdata += @(,("Datastore",($htmlsilver -bor $htmlitalics),$VMDisk.Filename.Substring($VMDisk.Filename.IndexOf("[")+1,$VMDisk.Filename.IndexOf("]")-1),$htmlwhite))
            $rowdata += @(,("Disk Path",($htmlsilver -bor $htmlitalics),$VMDisk.Filename.Substring($VMDisk.Filename.IndexOf("]")+2),$htmlwhite))
            $rowdata += @(,("Format",($htmlsilver -bor $htmlitalics),$VMDisk.StorageFormat,$htmlwhite))
            $rowdata += @(,("Type",($htmlsilver -bor $htmlitalics),$VMDisk.DiskType,$htmlwhite))
            $rowdata += @(,("Persistence",($htmlsilver -bor $htmlitalics),$VMDisk.Persistence,$htmlwhite))
        }
        If(($Snapshots) | Where {$_.VM -like $VM.Name}){$rowdata += @(,("VM has Snapshots",($htmlsilver -bor $htmlbold),(($Snapshots) | Where {$_.VM -like $VM.Name}).Count,$htmlwhite))}

        FormatHTMLTable "VM: $($VM.Name)" -rowArray $rowdata -noHeadCols 2 -fixedWidth $colWidths -tablewidth "350"
        WriteHTMLLine 0 0 " "
    }
}

#endregion

#region vCenter Issues functions

Function ProcessSnapIssues
{
    If($Snapshots)
    {
        Write-Verbose "$(Get-Date): Processing Issues: Virtual Machine Snapshots found"
        If($MSWord -or $PDF)
	    {
		    $Selection.InsertNewPage()
		    WriteWordLine 1 0 "Virtual Machines with Snapshots"
	    }
	    ElseIf($Text){Line 0 "Issue: Virtual Machines with Snapshots"}
            
        OutputSnapIssues $Snapshots

    }
}

Function OutputSnapIssues
{
    Param([object] $VMSnaps)

    If($MSWord -or $PDF)
    {

    }
    ElseIf($HTML)
    {
        WriteHTMLLine 0 0 " "
        $rowdata = @()
        $columnHeaders = @("Virtual Machine",($htmlsilver -bor $htmlbold),"Snapshot Name",($htmlsilver -bor $htmlbold),"Created",($htmlsilver -bor $htmlbold),"Running Current",($htmlsilver -bor $htmlbold),"Parent",($htmlsilver -bor $htmlbold),"Quiesced",($htmlsilver -bor $htmlbold),"Description",($htmlsilver -bor $htmlbold))

        foreach($Snap in $VMSnaps)
        {
            $rowdata += @(,($Snap.VM,$htmlwhite,$Snap.Name,$htmlwhite,$Snap.Created,$htmlwhite,$Snap.IsCurrent,$htmlwhite,$Snap.ParentSnapshot,$htmlwhite,$Snap.Quiesced,$htmlwhite,$Snap.Description,$htmlwhite))
        }

        FormatHTMLTable "VMs with Snapshots" -rowArray $rowdata -columnArray $columnHeaders
    }
    ElseIf($Text)
    {

    }
}

Function ProcessOpticalIssues
{
    $VMCDRom = $VirtualMachines | Where {$_.CDDrives.ConnectionState.Connected}
    If($VMCDRom)
    {
        Write-Verbose "$(Get-Date): Processing Issues: Mounted CDROM drives found"
        If($MSWord -or $PDF)
	    {
		    $Selection.InsertNewPage()
		    WriteWordLine 1 0 "Virtual Machines with CDROM drives mounted"
	    }
	    ElseIf($Text){Line 0 "Virtual Machines with CDROM drives mounted"}

        OutputOpticalIssues $VMCDRom
    }
    
}

Function OutputOpticalIssues
{
    Param([object] $VMCDRoms)

    If($MSWord -or $PDF)
    {

    }
    ElseIf($HTML)
    {
        WriteHTMLLine 0 0 " "
        $rowdata = @()
        $columnHeaders = @("Virtual Machine",($htmlsilver -bor $htmlbold),"ISO Path",($htmlsilver -bor $htmlbold),"Host Device",($htmlsilver -bor $htmlbold),"Remote Device",($htmlsilver -bor $htmlbold))

        foreach($VMCD in $VMCDRoms)
        {
            $rowdata += @(,($VMCD.Name,$htmlwhite,$VMCD.CDDrives.IsoPath,$htmlwhite,$VMCD.CDDrives.HostDevice,$htmlwhite,$VMCD.CDDrives.RemoteDevice,$htmlwhite))
        }

        FormatHTMLTable "VMs with mounted CDROM" -rowArray $rowdata -columnArray $columnHeaders
    }
    ElseIf($Text)
    {

    }
}

#endregion

#region script setup function
Function ProcessScriptSetup
{
	$script:startTime = Get-Date

	#$ComputerName = TestComputerName $ComputerName
}
#endregion

#region script end function
Function ProcessScriptEnd
{
	Write-Verbose "$(Get-Date): Script has completed"
	Write-Verbose "$(Get-Date): "

	#http://poshtips.com/measuring-elapsed-time-in-powershell/
	Write-Verbose "$(Get-Date): Script started: $($Script:StartTime)"
	Write-Verbose "$(Get-Date): Script ended: $(Get-Date)"
	$runtime = $(Get-Date) - $Script:StartTime
	$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds", `
		$runtime.Days, `
		$runtime.Hours, `
		$runtime.Minutes, `
		$runtime.Seconds,
		$runtime.Milliseconds)
	Write-Verbose "$(Get-Date): Elapsed time: $($Str)"

	If($Dev)
	{
		If($SmtpServer -eq "")
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
		}
		Else
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
		}
	}

	If($ScriptInfo)
	{
		$SIFile = "$($pwd.Path)\VMwareInventoryScriptInfo_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
		Out-File -FilePath $SIFile -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Add DateTime  : $($AddDateTime)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Chart         : $($Chart)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Company Name  : $($Script:CoName)" 4>$Null		
		}
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Cover Page    : $($CoverPage)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Dev           : $($Dev)" 4>$Null
		If($Dev)
		{
			Out-File -FilePath $SIFile -Append -InputObject "DevErrorFile  : $($Script:DevErrorFile)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Export        : $($Export)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Filename1     : $($Script:FileName1)" 4>$Null
		If($PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Filename2     : $($Script:FileName2)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Folder        : $($Folder)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "From          : $($From)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Full          : $($Full)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Import        : $($Import)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Issues        : $($Issues)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As HTML  : $($HTML)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As PDF   : $($PDF)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As TEXT  : $($TEXT)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As WORD  : $($MSWORD)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script Info   : $($ScriptInfo)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Port     : $($SmtpPort)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Server   : $($SmtpServer)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Title         : $($Script:Title)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "To            : $($To)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Use SSL       : $($UseSSL)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "User Name     : $($UserName)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "VIServerName  : $($VIServerName)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "OS Detected   : $($Script:RunningOS)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PoSH version  : $($Host.Version)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSCulture     : $($PSCulture)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSUICulture   : $($PSUICulture)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Word language : $($Script:WordLanguageValue)" 4>$Null
			Out-File -FilePath $SIFile -Append -InputObject "Word version  : $($Script:WordProduct)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script start  : $($Script:StartTime)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Elapsed time  : $($Str)" 4>$Null
	}

	$ErrorActionPreference = $SaveEAPreference
}
#endregion

#region email function
Function SendEmail
{
	Param([string]$Attachments, [string]$Subject)
	Write-Verbose "$(Get-Date): Prepare to email"
	$emailAttachment = $Attachments
	$emailSubject = $Script:Title
	$emailBody = @"
Hello, <br />
<br />
$Script:Title is attached.
"@ 

	If($Dev)
	{
		Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
	}
	
	$error.Clear()
	
	If($UseSSL)
	{
		Write-Verbose "$(Get-Date): Trying to send email using current user's credentials with SSL"
		Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
		-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
		-UseSSL *>$Null
	}
	Else
	{
		Write-Verbose  "$(Get-Date): Trying to send email using current user's credentials without SSL"
		Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
		-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To *>$Null
	}

	$e = $error[0]

	If($e.Exception.ToString().Contains("5.7.57"))
	{
		#The server response was: 5.7.57 SMTP; Client was not authenticated to send anonymous mail during MAIL FROM
		Write-Verbose "$(Get-Date): Current user's credentials failed. Ask for usable credentials."

		If($Dev)
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
		}

		$error.Clear()

		$emailCredentials = Get-Credential -Message "Enter the email account and password to send email"

		If($UseSSL)
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL -credential $emailCredentials *>$Null 
		}
		Else
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-credential $emailCredentials *>$Null 
		}

		$e = $error[0]

		If($? -and $Null -eq $e)
		{
			Write-Verbose "$(Get-Date): Email successfully sent using new credentials"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Email was not sent:"
			Write-Warning "$(Get-Date): Exception: $e.Exception" 
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): Email was not sent:"
		Write-Warning "$(Get-Date): Exception: $e.Exception" 
	}
}
#endregion

#region script core
#Script begins

ProcessScriptSetup

If(!($Import)){VISetup $VIServerName}

SetGlobals

SetFileName1andFileName2 "$($VIServerName)-Inventory"
[string]$Script:Title = "VMware Inventory Report - $VIServerName"

If($Issues)
{
    ProcessSummary
    ProcessSnapIssues
    ProcessOpticalIssues
}
Else
{
    ProcessSummary
    ProcessvCenter
    ProcessClusters
    ProcessResourcePools
    ProcessVMHosts
}

#Process full inventory
If($Full)
{
    ProcessDatastores
    ProcessHostNetworking
    ProcessStandardVSwitch
    ProcessVMKPorts
    ProcessVMPortGroups
    ProcessVirtualMachines
}

#endregion

#region finish script
Write-Verbose "$(Get-Date): Finishing up document"

#Disconnect from VCenter
If(!($Import)){Disconnect-VIServer $VIServerName -Confirm:$False 4>$Null}

#end of document processing

###Change the two lines below for your script###
$AbstractTitle = "VMware Inventory Report"
$SubjectTitle = "VMware vCenter Inventory Report"

UpdateDocumentProperties $AbstractTitle $SubjectTitle

ProcessDocumentOutput

ProcessScriptEnd
#endregion