<#
.SYNOPSIS
Script to automate the Department or OpCo Site configuration
.DESCRIPTION
It automates the process of group, list and  libraries creation and sets the current navigation. 
.NOTES
Copyright Infosys Limited, All rights reserved.
#>

#assigfn proper url inquick launch on rollback
# Parameters to get input from the user

#Param(

# [Parameter(Mandatory=$true)]
#	[ValidateScript({
#      If ($_ -ne '') {
#        $True
#      }
#      else {
#        Throw "$_ is empty"
#      }
#    })]
# [string]$UserName,

# [Parameter(Mandatory=$true)]
#	[ValidateScript({
#      If ($_.Length -ne 0) {
#        $True
#      }
#      else {
#        Throw "$_ is empty"
#      }
#    })]
# [Security.SecureString]${Password},

# [Parameter(Mandatory=$true)]
#	[ValidateScript({
#      If ($_ -ne '') {
#        $True
#      }
#      else {
#        Throw "$_ is empty"
#      }
#    })]
# [string]$SiteUrl,

# [Parameter(Mandatory=$true)]
#	[ValidateScript({
#      If ($_ -ne '') {
#        $True
#      }
#      else {
#        Throw "$_ is empty"
#      }
#    })]
# [string]$SiteName,
# [Parameter(Mandatory=$true,HelpMessage="Path to ...")]
#	[ValidateScript({
#      If ($_ -ne '') {
#        $True
#      }
#      else {
#        Throw "$_ is empty"
#      }
#    })]
# [string]$SiteShortName
# )

 #$global:SiteUrl="https://sysco.sharepoint.com/sites/OpCo_devtwo/";
 #$global:UserName="vmen8599@corp.sysco.com"
 #$Password ="Sysco123" | ConvertTo-SecureString -AsPlainText -Force
 #$SiteName ="Baraboo"
 #$SiteShortName="018-Baraboo"

 $global:SiteUrl="https://kabali11.sharepoint.com/sites/test1/";
 $global:UserName="admin@kabali11.onmicrosoft.com"
 $Password ="Infy@123" | ConvertTo-SecureString -AsPlainText -Force
 $SiteName ="Riverside"
 $SiteShortName="320-Riverside"


 


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

 $spContext = New-Object Microsoft.SharePoint.Client.ClientContext($Url)

 if($Url.Contains(".sharepoint.com")) # SharePoint Online
 {	
	$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $securePassword)
 }

 #See if we can establish a connection
 $spContext.Credentials = $credentials
 $spContext.RequestTimeOut = 5000 * 60 * 10;
 $web = $spContext.Web
 $site = $spContext.Site
 $spContext.Load($web)
 $spContext.Load($site)
 try
 {
	$spContext.ExecuteQuery()
	Log "Established connection to SharePoint at $Url OK" Green
}
catch
{
	Log "Not able to connect to SharePoint at $Url. Exception:$_.Exception.Message" red
	EndStatus
}

 return $spContext
}

 function Set-SharePointGroup($spContext){
	 <#
        .DESCRIPTION
        creates the custom SharePoint groups based on site type and assigns it to Sharepoint Site .

        .PARAMETER spContext
        Specifies the Sharepoint Context.

        #>
$web = $spContext.Web
$spContext.Load($web)
$spContext.Load($web.SiteGroups)
	 $web.Url
	 $spContext.Load($web.RoleDefinitions)
if($RollBack -ne "true"){
Log "Setting SharePoint Groups to SharePoint Online site $SiteUrl" Yellow

foreach($groupName in $GroupNames)
{

#check whether group exists

$groupExists=$web.SiteGroups.GetByName($groupName);
$spContext.Load($groupExists);
try
{
$spContext.ExecuteQuery()
$newGroup=$groupExists;
}
catch
{
# Create new Sharepoint Group object 
$newGroupInfo = New-Object Microsoft.SharePoint.Client.GroupCreationInformation  
$newGroupInfo.Title =   $groupName
$newGroupInfo.Description = $groupName  
$newGroup = $web.SiteGroups.Add($newGroupInfo) 
$newGroup.OnlyAllowMembersViewMembership=$false
$newGroup.update()
$spContext.Load($newGroup) 
} 
if($groupName -eq 'CommMaintenanceAdmin')
{

# Add current user to 'CommMaintenanceAdmin' Group
$userInfo = $web.EnsureUser($UserName)  
$spContext.Load($userInfo)  
$addUser = $newGroup.Users.AddUser($userInfo)  
$spContext.Load($addUser) 
$access = $web.RoleDefinitions.GetByName("Full Control")  
$groupOwner=$newGroup;
}
else
{
$access = $web.RoleDefinitions.GetByName($PermissionName)  
$groupOwner=$web.SiteGroups.GetByName("CommMaintenanceAdmin")
}

# Avoid adding OpCoDocLibraryUsers to Site Permission 

if($groupName -ne "OpCoDocLibraryUsers")
{
$roleAssignment =  New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($spContext)  
$roleAssignment.Add($access)  
$addPermission = $web.RoleAssignments.Add($newGroup, $roleAssignment)   

$spContext.Load($groupOwner)  
$spContext.Load($web)  
$spContext.Load($addPermission) 
$newGroup.Owner= $groupOwner 
$newGroup.update()
$web.Update()  
}
try
 {
	$spContext.ExecuteQuery()
	Log "Added SharePoint Group - $groupName -OK" Green
}
catch
{
	Log "Not able to add Sharepoint Group - $groupName. Exception:$_.Exception.Message" red
	EndStatus
}  
}
}
}

 function Remove-SharePointGroup($spContext){
	  <#
        .DESCRIPTION
        Removes the respecive SharePoint groups based on site type.

        .PARAMETER spContext
        Specifies the Sharepoint Context.

        #>
	 if($RollBack -ne "true")
	 {
	 Log "Removing OOTB SharePoint Groups from Site Permissions at $SiteUrl" Yellow
	 }
	 else
	 {
	 Log "Removing custom SharePoint Groups at $SiteUrl" Yellow
	 }

	 $web = $spContext.Web
	 $spContext.Load($web)
	 $groups = $web.SiteGroups # Gets all site groups 
	 $spContext.Load($groups) 
	 $owner= $web.AssociatedOwnerGroup
	 $member= $web.AssociatedMemberGroup
	 $visitor= $web.AssociatedVisitorGroup
	 $spContext.Load($owner)
	 $spContext.Load($member)
	 $spContext.Load($visitor)
	 try
 {
	$spContext.ExecuteQuery() 
	for ($i=0; $i -lt $groups.Count; $i++){
		if($GroupNames -match $groups[$i].Title)
		{
			if($RollBack -eq "true")
			{
				$groupExists= $web.SiteGroups.GetById($groups[$i].ID)
				$spContext.Load($groupExists)
				try
				{
					# Delete custom SharePoint Groups
					$spContext.ExecuteQuery()
					if($groupExists)
					{
					$web.SiteGroups.RemoveById($groups[$i].ID)
					}
				}
				catch
				{
				Log "Group does not exists - $groupName. Exception:$_.Exception.Message" red
				}
				
			}
			else
			{
			#do nothing
			}
		}
		else
		{
			# Remove OOTB SP Groups
			if($RollBack -ne "true" -and $owner.Title -ne $groups[$i].Title -and $member.Title -ne $groups[$i].Title -and $visitor.Title -ne $groups[$i].Title  )
			{
			$web.RoleAssignments.Groups.Remove($groups[$i])
			}
		}
}
$spContext.ExecuteQuery()
Log "Removed respective SharePoint Group -OK" Green
}
catch
{
	Log "Not able to remove Sharepoint Group - $groupName. Exception:$_.Exception.Message" red
	EndStatus
}  
	}

 function Set-QuickLaunch($spContext){
	   <#
        .DESCRIPTION
        Sets the Quick Launch of the site.

        .PARAMETER spContext
        Specifies the Sharepoint Context.

        #>
		Log "Setting QuickLaunch Navigation at $SiteUrl" Yellow
		$web = $spContext.Web
		$spContext.Load($web)
		$quickLaunchColl= $web.Navigation.QuickLaunch
		$spContext.Load($quickLaunchColl)
		$spContext.ExecuteQuery()
		# Iterate and delete the OOTB Navigation nodes
		for ($i = $quickLaunchColl.Count-1; $i -ge 0; $i--)
		{
		$quickLaunchColl[$i].DeleteObject()
		}

		$web.update()
		$spContext.Load($quickLaunchColl)  
		try
		{
		$spContext.ExecuteQuery()
		Log "Junk Current navigation nodes are removed -OK" Green
		# Set Quick launch nodes 
		
		foreach($link in $Navigation)
		{
		$navColl = $web.Navigation.QuickLaunch
		$node = $navColl | where { $_.Title -eq $link.Header }
		if($link.Header -ne "Empty" -and !$node)
		{
		$node = New-Object Microsoft.SharePoint.Client.NavigationNodeCreationInformation
		$node.Title = $link.Header
		$node.Url="#"
		$node.AsLastNode = $true
		$spContext.Load($navColl.Add($node))
		}
		$newNavNode = New-Object Microsoft.SharePoint.Client.NavigationNodeCreationInformation
		$newNavNode.Title = $link.Name
			if($RollBack -ne "true")
			{
		$newNavNode.Url = $RootUrl.TrimEnd('/')+"/" +$link.URL
			}
			else
			{
		$newNavNode.Url = $SiteUrl.TrimEnd('/')+"/" +$link.URL
			}
		$newNavNode.AsLastNode = $true
		if($link.Header -ne "Empty")
		{
		$parentNode = $navColl | where { $_.Title -eq $link.Header }
		$spContext.Load($parentNode.Children.Add($newNavNode))
		}
		else
		{
		$spContext.Load($navColl.Add($newNavNode))
		}
		}    
		$spContext.ExecuteQuery()
	    Log "Current navigation nodes are set -OK" Green
		}
		catch
		{
	    Log "Error in Setting the Current Site navigation - $SiteUrl. Exception:$_.Exception.Message" red
	    EndStatus
		}
	 }

 function Readvaluesfromconfig($BaseDirectory,$BaseConfig,$SiteType){
	  <#
        .DESCRIPTION
        Reads the configuration values from JSON in string format.

        .PARAMETER BaseDirectory
        Specifies the directory path.

		.PARAMETER BaseConfig
        Specifies the file path.

		.PARAMETER SiteType
        Specifies the Site type whether Department or OpCo.
	 
        #>
		 
# Load and parse the JSON configuration file
try {
	$Config = Get-Content "$BaseDirectory$BaseConfig" -Raw -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue | ConvertFrom-Json -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue
} catch {
	Log "The Base configuration file is missing!" red
	EndStatus
}
	 # Check the configuration
if (!($Config)) {
	Log "The Base configuration file is missing!" red
}
else
{
	if($SiteType -eq "Department")
	{
	$global:GroupNames = ($Config.Department.GroupNames)
	$global:Navigation =($Config.Department.Navigation)
	$global:Libraries =($Config.Department.Libraries)
	}
	else
	{
	$global:GroupNames = ($Config.OpCo.GroupNames)
	$global:Navigation =($Config.OpCo.Navigation)
	$global:Libraries =($Config.OpCo.Libraries)
	}
	$global:LogFlag=($Config.Logging.enabled)
	$global:RootUrl=($Config.Configuration.RootSiteURL)
	$global:PermissionName=($Config.Configuration.PermissionLevel)

	# Get user input for Deploy Mechanism
	$caption = ""    
	$message = "PLEASE CONFIRM YOUR OPERATION METHOD (DEPLOY\ROLLBACK)"
	[int]$defaultChoice = 0
	$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Deploy", "Deploy"
	$no = New-Object System.Management.Automation.Host.ChoiceDescription "&Rollback", "Rollback"
	$exit=New-Object System.Management.Automation.Host.ChoiceDescription "&Exit", "Exit"
	$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no,$exit)
	$choiceRTN = $host.ui.PromptForChoice($caption,$message, $options,$defaultChoice)
	if ( $choiceRTN -eq 0 )
	{
	$global:RollBack = "false"
	}
	elseif ( $choiceRTN -eq 1 )
	{
	$global:RollBack = "true"
	}
	else
	{
	Log "Exiting the Powershell command" red
	EndStatus
	}
	
	if($RollBack -eq "true")
	{
	$global:Navigation =($Config.RollBack.Navigation)
	}
	$LogTime = Get-Date -Format "MM-dd-yyyy_hh-mm-ss"
	Log "Logging Started at $LogTime" White
	Log "Configuration Values are parsed successfully" Green
}
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
	if($LogFlag -eq 'true')
	{$string | out-file -Filepath $logfile -append}
}

 function Create-Lists($spContext){ 
	   <#
        .DESCRIPTION
        Creates Lists/Libraries based on the Site Type Chosen.

		.PARAMETER spContext
        Specifies the Sharepoint Context.

        #>
	
	if($RollBack -ne "true")
	{
	Log "Creating libraries/list at  $SiteUrl" Yellow
	}
	else
	{
	Log "Deleting libraries/list at  $SiteUrl" Yellow
	}
	$web = $spContext.Web
	$spContext.Load($web)
	$spContext.Load($web.SiteGroups)
	$spContext.Load($web.Lists)
	$spContext.ExecuteQuery()
	 foreach($lib in $libraries)
	 {
		$LibraryName =$lib.DisplayName
		# Check if Lists Exists
		$listExists = $web.Lists | where{$_.Title -eq $lib.DisplayName}
		if(!$listExists -and $RollBack -ne "true")
		{
		try{
		# New List Creation object 
		$ListInfo = New-Object Microsoft.SharePoint.Client.ListCreationInformation
		$ListInfo.Title = $lib.InternalName
		$ListInfo.Description=$lib.Description
		$ListInfo.TemplateType = [Microsoft.SharePoint.Client.ListTemplateType] $lib.Type
		$List = $web.Lists.Add($ListInfo)
		$List.OnQuickLaunch=$false;
		$List.Update()
		# Add field XML
		if($lib.Flag -ne "Folders")
		{
		foreach($field in $lib.Fields)
		{
		$List.Fields.AddFieldAsXml($field,$true,[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
		}
		}
		else
		{
		foreach($name in $lib.Folders)
		{
		$folder=$List.RootFolder.Folders.Add($name);
        $spContext.Load($folder);
		}
		}
		$spContext.Load($List)
		$spContext.ExecuteQuery()

		#update List and columns	
		$List.Title=$lib.DisplayName
		foreach($field in $lib.Fields)
		 {
		$xml=[xml]$Field
		$UpdateValue=$xml.Field | Select DisplayName,Name
		$FieldtobeUpdated=$List.Fields.GetByInternalNameOrTitle($UpdateValue.DisplayName)
		$FieldtobeUpdated.Title = $UpdateValue.Name
		$FieldtobeUpdated.Update()
		}
		if($lib.Type -ne "DocumentLibrary")
		{
		$titleField=$List.Fields.GetByInternalNameOrTitle("Title")
		$titleField.Required=$false
		$titleField.SetShowInDisplayForm($false)
		$titleField.SetShowInEditForm($false)
		$titleField.SetShowInNewForm($false)
		$titleField.Hidden=$true;
		$titleField.update()
		$listView=$List.DefaultView
		$listView.ViewFields.Remove("LinkTitle")
		$List.OnQuickLaunch=$lib.Navigation
		$List.NoCrawl = $true
		$listView.update()
		}
		else
		{
		$List.OnQuickLaunch=$lib.Navigation
		$List.EnableFolderCreation=$lib.NewFolder
		$List.ForceCheckout=$lib.CheckOut
		$List.EnableVersioning=$lib.EnableVersion;
		$List.NoCrawl = $true
		}
		$List.update()
		$userPermission= $web.EnsureUser($UserName)  
		$spContext.Load($userPermission)
		$spContext.ExecuteQuery()
		
		# Set List permissions
		$List.BreakRoleInheritance($false,$false)

		foreach($groupName in $GroupNames)
	    {
			if($groupName -ne "OpCoDocLibraryUsers" -and $lib.Flag -ne "Folders")
			{
			$groupPermission=$web.SiteGroups.GetByName($groupName);
			}
			elseif($lib.Flag -eq "Folders" -and $groupName -ne "CommOpCoUsers")
			{
			$groupPermission=$web.SiteGroups.GetByName($groupName);
			}
			else
			{
			$groupPermission=''
			}
			if($groupPermission -ne '')
			{
			if($groupName -eq 'CommMaintenanceAdmin')
			{
			$access = $web.RoleDefinitions.GetByName("Full Control")  
			}
			else
			{
			$access = $web.RoleDefinitions.GetByName("Contribute")  
			}
			$roleAssignment =  New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($spContext)  
			$roleAssignment.Add($access)  
			$addPermission = $List.RoleAssignments.Add($groupPermission, $roleAssignment)   
			$spContext.Load($List)  
			$spContext.Load($addPermission)  
			}
		}
		$userroleassignment=$List.RoleAssignments.GetByPrincipalId($userPermission.Id);
		$userroleassignment.RoleDefinitionBindings.RemoveAll()
		$userroleassignment.Update()
		$spContext.ExecuteQuery()
		Log "Library/List is created and updated $LibraryName -OK" Green
		}
		catch
		{
		Log "Error in Creating List(s) $LibraryName - $SiteUrl. Exception:$_.Exception.Message" red
		EndStatus
		}
		}
		else
		{
			# Rollback Mechanism to delete custom lists
			if($RollBack -ne "true")
			{
				Log "Warning ! List already exists $LibraryName - $SiteUrl" red
			}
			elseif($RollBack -eq "true" -and $listExists)
			{
				try{
				$listDelete=$web.Lists.GetByTitle($lib.DisplayName)
				$spContext.Load($listDelete)
				$listDelete.DeleteObject()
				$spContext.ExecuteQuery()
				Log "Deleting custom list $LibraryName - $SiteUrl" Green
			}
				catch
				{
				Log "Error in Deleting List $LibraryName - $SiteUrl. Exception:$_.Exception.Message" red
				EndStatus
				}
			}
			else
			{
			 Log "No List -$LibraryName Exists at -$SiteUrl" Green
			}
		 }
}
	
}

 function Create-PermissionLevel($spContext){ 
<#
        .DESCRIPTION
        Creates custom permission level.

		.PARAMETER spContext
        Specifies the Sharepoint Context.

#>
	$permName=$PermissionName
	$permDescription="Sysco Custom Level Permission created by PowerShell"
	$permissionString = "ViewListItems,OpenItems,ViewFormPages,ViewPages,BrowseUserInfo,UseRemoteAPIs,Open"

	$web=$spContext.Web
	$spContext.Load($web)
	$roleDefinitionCol = $web.RoleDefinitions
    $spContext.Load($roleDefinitionCol)
    $spContext.ExecuteQuery()
    $permExists = $false

    #Check if the permission level is exists or not
    foreach($role in $roleDefinitionCol)
    {
       if($role.Name -eq $permName)
        {
            $permExists = $True
        }
    }
      if($RollBack -ne "true")
	 {
    Log "Creating Pemission level with the name $permName" Yellow
	}
	 else
	 {
	Log "Deleting Pemission level with the name $permName" Yellow
	 }
    if($permExists -ne $True -and $RollBack -ne "true")
    {
        try
        {
            $spRoleDef = New-Object Microsoft.SharePoint.Client.RoleDefinitionCreationInformation
            $spBasePerm = New-Object Microsoft.SharePoint.Client.BasePermissions
                
            $permissions = $permissionString.split(",");
            foreach($perm in $permissions)
            {
                $spBasePerm.Set($perm) 
            }

            $spRoleDef.Name = $permName
            $spRoleDef.Description = $permDescription
            $spRoleDef.BasePermissions = $spBasePerm    

            $roleDefinition = $web.RoleDefinitions.Add($spRoleDef)

            $spContext.ExecuteQuery()

            Log "Pemission level with the name $permName created" Green
        }
        catch
        {
            Log "There was an error creating Permission Level $permName : Error details $_.Exception.Message" Red
			EndStatus
		}
    }
    else
    {
		if($RollBack -ne "true")
		{
        Log "Pemission level with the name $permName already exists" Red
		}
		elseif($permExists -eq $True)
		{
			
			$removePermissionLevel=$web.RoleDefinitions.GetByName($PermissionName)
			$removePermissionLevel.DeleteObject()
			try{
			$spContext.ExecuteQuery()
				Log "Pemission level with the name $permName deleted" Green
				}
			catch
			{
			Log "There was an error in deleting Permission Level $permName : Error details $_.Exception.Message" Red
			EndStatus
			}
		}
		else
		{
		Log "Pemission level with the name $permName does not exist" Yellow
		}
    }
}

 function Check-Feature($spContext){ 
<#
        .DESCRIPTION
        Checks whether Publishing feature is enabled in the site.

		.PARAMETER spContext
        Specifies the Sharepoint Context.

#>
Log "Checking for Feature Activation" Yellow
$site = $spContext.Site
$spContext.Load($site.Features)
$featureguid = new-object System.Guid "f6924d36-2fa8-4f0b-b16d-06b7250180fa"  
$feature = $site.Features.GetById($featureguid)
$spContext.Load($feature)
	try{
$spContext.ExecuteQuery()
if ($feature.DefinitionId -eq $null)
	{ 
	Log "Activating Publishing Feature at site level - $SiteUrl" Yellow
	$site.Features.Add($featureguid, $true, [Microsoft.SharePoint.Client.FeatureDefinitionScope]::None);  
	$spContext.ExecuteQuery()
	Log "Feature Activated - OK" green
	} 
	else
	 { 
		Log "Checked for feature activation - OK" green
	 }
	
	}
	catch
	{
	Log "Error in checking the feature activation - $SiteUrl. Exception:$_.Exception.Message" red
	EndStatus
	
	}
	}

 function Create-Item($itemContext,$ListTitle,$Name,$ShortName,$Random){
	 	  <#
        .DESCRIPTION
        Funtion to create an item.

		.PARAMETER itemContext
        Specifies the Sharepoint Context.

		.PARAMETER ListTitle
        Specifies the List Title.

		.PARAMETER Name
        Specifies the name of the site.

		.PARAMETER ShortName
        Specifies the Short name of the site.

		.PARAMETER Random
        Specifies the generated random number.
#>
	 Item-Exists $itemContext $ListTitle
	 if ($items.Count -eq 0) 
	 { 
	 $List = $itemContext.Web.Lists.GetByTitle($ListTitle)
	 $ListItemCreationInformation = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
	 $NewListItem = $List.AddItem($ListItemCreationInformation)
	 if($ListTitle -eq "Document Types")
	 {
	 $NewListItem["Title"] =$Random
	 $NewListItem["DocTypes"]="Data Entry;Pay Files;ACH/Wire;Personnel Change Form;Mass Upload;OpCo Authorizations;OpCO Reporting;Child Support/Garnishment"
	 $NewListItem["URL"]=$RelativeURL
	 }
	else
	{
	$NewListItem["Title"] =$Random
	$NewListItem[$Name]=$SiteName
	if($ListTitle -eq "Opco Roles")
	{
	$NewListItem[$ShortName]="Division HRBP;HR Admin Assistant;HR Advisor;HR Business Partner - Colossal Jumbo;HR Business Partner - Large Medium;HR Business Partner - Super Colossal;HR Coordinator;HR Generalist;HR Supervisor;Human Resources Director;Human Resources Manager;Human Resources VP;Lead HR Business Partner;Sr HR Manager"
	}
	else
	{
	$NewListItem[$ShortName]=$SiteShortName
	}
	$NewListItem["URL"]=$RelativeURL
	}
	 $NewListItem.Update()
 }
}

 function Item-Exists($itemContext,$ListTitle){
	<#
        .DESCRIPTION
        Funtion to delete an item.

		.PARAMETER itemContext
        Specifies the Sharepoint Context.

		.PARAMETER ListTitle
        Specifies the List Title.
#>
	$List = $itemContext.Web.Lists.GetByTitle($ListTitle)      
	$spQuery = New-Object Microsoft.SharePoint.Client.CamlQuery 
	 $spQuery.ViewXml=  "<View><Query><Where><Eq><FieldRef Name='URL' /><Value Type='Text'>$RelativeURL</Value></Eq></Where></Query></View>"
    $global:items = $List.GetItems($spQuery)
	$itemContext.Load($items) 
	 try
	 {
	$itemContext.ExecuteQuery()
	 } 
	 catch
	 {
	Log "Error in querying items at $RootUrl. Exception:$_.Exception.Message" red
	EndStatus
	 } 
 }

 function Delete-Item($itemContext,$ListTitle){
	  <#
        .DESCRIPTION
        Funtion to delete an item.

		.PARAMETER itemContext
        Specifies the Sharepoint Context.

		.PARAMETER ListTitle
        Specifies the List Title.
#>
	 Item-Exists $itemContext $ListTitle
	 if ($items.Count-gt0)  
    {  
        for ($i=$items.Count-1; $i -ge 0; $i--)  
        {  
            $items[$i].DeleteObject()  
        }   
    }  

 }

 function Create-Entry($itemContext,$SiteName,$SiteShortName,$RelativeURL,$SiteType){
	 <#
        .DESCRIPTION
        Funtion which defines to create or delete an item.

		.PARAMETER itemContext
        Specifies the Sharepoint Context.

		.PARAMETER SiteName
        Specifies the Site name.

		.PARAMETER SiteShortName
        Specifies the Site Shortname.

		.PARAMETER RelativeURL
        Specifies the Site RelativeURL.

		.PARAMETER SiteType
        Specifies the SiteType (Department/OpCo).

#>
	 
	 $rnd= Get-Random -Maximum 9999 -Minimum 0001
	 if($SiteType -eq "Department")
	 {	
		 if($RollBack -ne "true")
	 {
		Log "Adding Entry in Departments & Document Types List at $RootUrl" Yellow
		Create-Item $itemContext "Departments" "DepartmentName" "DeptShortName" "HR$rnd"
		Create-Item $itemContext "Document Types" "" "" "HR$rnd"
		}
		else
		{
		Log "Deleting Entry in Departments & Document Types List at $RootUrl" Yellow
		Delete-Item $itemContext "Departments"
		Delete-Item $itemContext "Document Types"
		}
		
	 }
	 else
	 {
		if($RollBack -ne "true")
		{
		Log "Adding Entry in OpCos & OpCo Roles List at $RootUrl" Yellow
		Create-Item $itemContext "OpCos" "OpCoName" "OpCoShortName" "OpCo$rnd"
		Create-Item $itemContext "Opco Roles" "OpCoShortName" "UserRoles" "OpCo$rnd"
		}
		else
		{
		Log "Deleting Entry in OpCos & OpCo Roles List at $RootUrl" Yellow
		Delete-Item $itemContext "OpCos"
		Delete-Item $itemContext "Opco Roles"
		}
	 }
try
	 {
     $itemContext.ExecuteQuery()
	 Log "Entries for new Site $RelativeURL has been added/Deleted - OK" green
	 }
	 catch
	 {
	Log "Error in creating entries - $RootUrl. Exception:$_.Exception.Message" red
	EndStatus
	 }
}

 function EndStatus(){
	  <#
        .DESCRIPTION
        Logs the time and exits the script flow.
        #>
	$LogTime = Get-Date -Format "MM-dd-yyyy_hh-mm-ss"
	Log "-------------------------- End - $LogTime --------------------------" Yellow
	exit 1
 }

#Ensure SharePoint PowerShell dll is loaded
 if((Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) -eq $null){
   Add-PSSnapin Microsoft.SharePoint.PowerShell
 }

#Add required Client Dlls 
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")

$BaseDirectory = split-path -parent $MyInvocation.MyCommand.Definition
$logfile = $BaseDirectory+"\LogFile.log"
$BaseConfig = "\config.json"

# user input for Site Type
$caption = ""    
$message = "PLEASE CONFIRM THE SITE CREATION TYPE"
[int]$defaultChoice = 0
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Department", "Department"
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&OpCo", "OpCo"
$exit=New-Object System.Management.Automation.Host.ChoiceDescription "&Exit", "Exit"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no,$exit)
$choiceRTN = $host.ui.PromptForChoice($caption,$message, $options,$defaultChoice)
if ( $choiceRTN -eq 0 )
{
 $SiteType="Department"
 Readvaluesfromconfig $BaseDirectory $BaseConfig $SiteType
}
elseif ( $choiceRTN -eq 1 )
{
 $SiteType="OpCo"
 Readvaluesfromconfig $BaseDirectory $BaseConfig $SiteType
}
else
{
Log "Exiting the Powershell command" red
EndStatus
}


$spContext = LoadAndConnectToSharePoint $SiteUrl $UserName $Password

$RelativeURL= $spContext.Web.ServerRelativeUrl

Check-Feature $spContext

#Create-PermissionLevel $spContext 

#Set-SharePointGroup $spContext

#Remove-SharePointGroup $spContext

#Create-Lists $spContext

#Set-QuickLaunch $spContext

#$itemContext = LoadAndConnectToSharePoint $RootUrl $UserName $Password

#Create-Entry $itemContext $SiteName $SiteShortName $RelativeURL $SiteType

EndStatus

# End


