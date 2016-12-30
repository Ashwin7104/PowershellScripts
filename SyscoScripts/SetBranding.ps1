#
# SetBranding.ps1


function EndStatus(){
	  <#
        .DESCRIPTION
        Logs the time and exits the script flow.
        #>
	$LogTime = Get-Date -Format "MM-dd-yyyy_hh-mm-ss"
	Log "-------------------------- End - $LogTime --------------------------" Yellow
	exit 1
 }

function Log($string, $color){
	   <#
        .DESCRIPTION
        Logs the progress to console and file.

        .PARAMETER string
        Specifies the message.

		.PARAMETER color
        Specifies the color in which the string are displayed in console.

        #>

   if ($Color -eq $null) {$color = "white"}
   write-host $string -foregroundcolor $color
	$string | out-file -Filepath $logfile -append
}

function Delete-Folders($Context,$ListName,$FileName){
		<#
        .DESCRIPTION
        Deletes the Branding Files.

        .PARAMETER Context
        Specifies the Sharepoint Context.

		.PARAMETER ListName
        Specifies the name of the library.

		.PARAMETER FileName
        Specifies the Branding File name.

        #>
	try
	{
	Log "Delete Sysco Branding Folder - Started" Yellow
	$List = $Context.Web.Lists.GetByTitle($ListName)
	$Context.Load($List.RootFolder.Folders)
	$Context.ExecuteQuery()
	for ($i=$List.RootFolder.Folders.Count-1; $i -ge 0; $i--){
	if($List.RootFolder.Folders[$i].name -eq $FileName)
	{
	$List.RootFolder.Folders[$i].DeleteObject()
	}
	}
	$Context.ExecuteQuery()
	Log "Delete Sysco Branding Folder - OK" Green
	}
	catch
	{
	Log "Not able to delete branding files at $SiteURL. Exception:$_.Exception.Message" red
	EndStatus
	}
}

function Delete-WelcomePage($Context,$ListName,$PageName){
	<#
        .DESCRIPTION
        Deletes the welcomepage of the site.

        .PARAMETER Context
        Specifies the Sharepoint Context.

		.PARAMETER ListName
        Specifies the name of the library.

		.PARAMETER PageName
        Specifies the page name.

        #>
	try
	{
	Log "Delete WelcomePage - Started" Yellow
	$List = $Context.Web.Lists.GetByTitle($ListName)
	$Context.Load($List.RootFolder.Files)
	$Context.ExecuteQuery()
	for ($i=$List.RootFolder.Files.Count-1; $i -ge 0; $i--){
	if($List.RootFolder.Files[$i].name -eq $PageName)
	{
	$List.RootFolder.Files[$i].DeleteObject()
	Log "Deleted WelcomePage - OK" Green
	}
	}
	$rootFolder = $Context.Web.RootFolder
	$rootFolder.WelcomePage = "SitePages/home.aspx"
	$rootFolder.Update()
	$Context.ExecuteQuery()
	}
	catch
	{
	Log "Not able to delete welcome page at $SiteURL. Exception:$_.Exception.Message" red
	EndStatus
	}
}

function Delete-CustomActions($Context){
	<#
        .DESCRIPTION
        Deletes the user custom actions from the site.

        .PARAMETER Context
        Specifies the Sharepoint Context.
        #>
	 try
	 {
	$web=$Context.Web
	$Context.Load($web)
	$Context.ExecuteQuery()
	 if($web)
	 {
		 $actions=$web.get_userCustomActions()
		 $Context.Load($actions)
		 $Context.ExecuteQuery()
		 if($actions)
		 {
			 Log "Deleting Existing actions" Yellow
			 for ($i=$actions.Count-1; $i -ge 0; $i--)
			 {
				Log $actions[$i].Description White
				if($actions[$i].Description -eq "JqueryFile" -or $actions[$i].Description -eq "SyscoBrandingCSS" -or $actions[$i].Description -eq "SyscoBrandingJS" )
				{
				$actions[$i].DeleteObject()
				}
			 }
		}
		$Context.ExecuteQuery()
	}
	}
	catch
	{
	Log "Not able to delete the existing custom actions $SiteURL. Exception:$_.Exception.Message" red
	EndStatus
	}
}

function Set-Link($Context){
<#
        .DESCRIPTION
        Injects the JS and CSS files to the masterpage.

        .PARAMETER Context
        Specifies the Sharepoint Context.
        #>
		try
		{
		$web=$Context.Web
		$Context.Load($web)
		$Context.ExecuteQuery()
		$URL=$web.Url
	    Log "Setting up the Branding" Yellow

		$customActionJS = $web.UserCustomActions.Add()
		$customActionJS.Description="JqueryFile"
		$customActionJS.Location = "ScriptLink";
		$customActionJS.ScriptSrc = $URL +"/Style Library/Sysco_BrandingFiles/jquery-1.11.2.min.js"
		$customActionJS.Sequence = 9;
		$customActionJS.Update();

		$customActionBranding = $web.UserCustomActions.Add()
		$customActionBranding.Description="SyscoBrandingJS"
		$customActionBranding.Location = "ScriptLink";
		$customActionBranding.ScriptSrc = $URL+"/Style Library/Sysco_BrandingFiles/Sysco_Branding.js"
		$customActionBranding.Sequence = 10;
		$customActionBranding.Update();

		$customActionCSS = $web.UserCustomActions.Add()
		$customActionCSS.Description="SyscoBrandingCSS"
		$customActionCSS.Location = "ScriptLink";
		$customActionCSS.ScriptBlock= "document.write('<link rel=""stylesheet"" type=""text/css"" href=""$URL/Style%20Library/Sysco_BrandingFiles/Sysco_Branding.css"" />');";
		$customActionCSS.Sequence = 11;
		$customActionCSS.Update();
		$Context.ExecuteQuery()

		Log "Setting up the Branding Completed" Green
		}
		catch
		{
		Log "Not able to set the Branding at $SiteURL. Exception:$_.Exception.Message" red
		EndStatus
		}
	}

function Upload-Files($Context,$DirFolder){
	  <#
        .DESCRIPTION
        Funtion to Upload the Brandind Files.

		.PARAMETER Context
        Specifies the Sharepoint Context.

		.PARAMETER DirFolder
        Specifies the path of the folder.
#>
	try{
Log "Upload Branding Files - Started" Yellow
$List = $Context.Web.Lists.GetByTitle('Style Library')
$Context.Load($List)
$Context.ExecuteQuery()

$curFolder = $List.RootFolder.Folders.Add("Sysco_BrandingFiles")
$Context.Load($curFolder)
$Context.ExecuteQuery()

#Upload file
Foreach ($File in (dir $DirFolder -File))
{
$FileStream = New-Object IO.FileStream($File.FullName,[System.IO.FileMode]::Open)
$FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
$FileCreationInfo.Overwrite = $true
$FileCreationInfo.ContentStream = $FileStream
$FileCreationInfo.URL = $File
$Upload = $curFolder.Files.Add($FileCreationInfo)
$Upload.CheckIn("Comments", 1)
$Context.Load($Upload)
$Context.ExecuteQuery()
}
Log "Upload Branding Files - OK" Green
		}
	catch
		{
		Log "Not able to upload branding files at $SiteURL. Exception:$_.Exception.Message" red
		EndStatus
		}
}

function Create-WikiPage([Microsoft.SharePoint.Client.ClientContext]$Context,[string]$ListName,[string]$PageName){
	<#
        .DESCRIPTION
        Creates the empty welcome page of the site.

        .PARAMETER Context
        Specifies the Sharepoint Context.

		.PARAMETER ListName
        Specifies the name of the library.

		.PARAMETER PageName
        Specifies the page name.

        #> 
	  try
	 {
	Log "Creating HomePage - /SitePages/welcome.aspx" Yellow
    $wikiLibrary = $Context.Web.Lists.GetByTitle($ListName)
    $Context.Load($wikiLibrary.RootFolder)
    $Context.ExecuteQuery()
    $wikiPageInfo = New-Object Microsoft.SharePoint.Client.Utilities.WikiPageCreationInformation
    $wikiPageInfo.WikiHtmlContent = ""
    $wikiPageInfo.ServerRelativeUrl = [String]::Format("{0}/{1}", $wikiLibrary.RootFolder.ServerRelativeUrl, $PageName)
    $wikiFile = [Microsoft.SharePoint.Client.Utilities.Utility]::CreateWikiPageInContextWeb($Context, $wikiPageInfo)
    $context.ExecuteQuery() 
	$rootFolder = $Context.Web.RootFolder
	$rootFolder.WelcomePage = "SitePages/Welcome.aspx"
	$rootFolder.Update()
	$context.ExecuteQuery() 
    Log "Homepage has been set - OK" Green
	}
	catch
	{
	Log "Not able to set the HomePage at $SiteURL. Exception:$_.Exception.Message" red
	EndStatus
	}
}

function LoadAndConnectToSharePoint($Url, $UserName, $securePassword) {
<#
        .DESCRIPTION
        Establishes the connection with O365 Sharepoint site and returns clientcontext.

        .PARAMETER URL
        Specifies the URL to establish the connection.

		.PARAMETER UserName
        Specifies the username to connect with the Site.

		.PARAMETER securePassword
        It holds the secure string password to access the site.

        #>

 Log "Establishing connection to SharePoint Online site $Url" Yellow

 $Context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)

 if($Url.Contains(".sharepoint.com")) # SharePoint Online
 {	
	$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $securePassword)
 }

 #See if we can establish a connection
 $Context.Credentials = $credentials
 $Context.RequestTimeOut = 5000 * 60 * 10;
 $web = $Context.Web
 $site = $Context.Site
 $Context.Load($web)
 $Context.Load($site)
 try
 {
	$Context.ExecuteQuery()
	Log "Established connection to SharePoint at $Url OK" Green
}
catch
{
	Log "Not able to connect to SharePoint at $Url. Exception:$_.Exception.Message" red
	EndStatus
}

 return $Context
}

#Ensure SharePoint PowerShell dll is loaded
 if((Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) -eq $null){
   Add-PSSnapin Microsoft.SharePoint.PowerShell
 }

#Add required Client Dlls 
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")

#$UserName = "lram8597@corp.sysco.com"
#$SiteURL = "https://sysco.sharepoint.com/sites/Baraboo_UAT"
#$Password ="Sysco123" | ConvertTo-SecureString -AsPlainText -Force

$UserName = "admin@kabali11.onmicrosoft.com"
$SiteURL = "https://kabali11.sharepoint.com/sites/ashwin"
$Password ="Infy@123" | ConvertTo-SecureString -AsPlainText -Force

#$User = Read-Host -Prompt "Enter the username"
#$Password = Read-Host -Prompt "Enter the password" -AsSecureString
#$SiteURL = Read-Host -Prompt "Enter the SIte URL"

$BaseDirectory = split-path -parent $MyInvocation.MyCommand.Definition
$DirFolder = "$BaseDirectory\Branding"
$logfile = $BaseDirectory+"\LogFile.log"
$LogTime = Get-Date -Format "MM-dd-yyyy_hh-mm-ss"

Log "Logging Started at $LogTime" White

# user input for Site Type
$caption = "" 
$type=''   
$message = "PLEASE CONFIRM THE DEPLOYMENT TYPE"
[int]$defaultChoice = 0
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Deploy", "Deploy"
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&Rollback", "Rollback"
$exit=New-Object System.Management.Automation.Host.ChoiceDescription "&Exit", "Exit"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no,$exit)
$choiceRTN = $host.ui.PromptForChoice($caption,$message, $options,$defaultChoice)
if ( $choiceRTN -eq 0 )
{
$type="Deploy"
}
elseif ( $choiceRTN -eq 1 )
{
$type="Rollback"
}
else
{
EndStatus
}
 
$Context = LoadAndConnectToSharePoint $SiteUrl $UserName $Password

Delete-CustomActions $Context

if($type -eq "Deploy")
{

Upload-Files $Context $DirFolder

Set-Link $Context

Create-WikiPage -Context $Context -ListName "Site Pages" -PageName "Welcome.aspx"

Log "Deployment Completed" Green
}
else
{

Delete-Folders $Context "Style Library" "Sysco_BrandingFiles"	

Delete-WelcomePage $Context "Site Pages" "Welcome.aspx"	
		 
Log "Rollback Completed" Green
}		


#
