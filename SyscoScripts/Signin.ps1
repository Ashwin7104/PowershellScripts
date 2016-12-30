#
# Signin.ps1
 [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
 [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")

$DeptUrl="https://kabali11.sharepoint.com/sites/test1";
 $UserName="admin@kabali11.onmicrosoft.com"
 $Password ="Infy@123" | ConvertTo-SecureString -AsPlainText -Force
 $spContext = New-Object Microsoft.SharePoint.Client.ClientContext($DeptUrl)
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $Password)


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
	 $ashwin= $web.AssociatedOwnerGroup
	 $spContext.Load($ashwin)
	 $spContext.ExecuteQuery()
	 Write-Host $ashwin.Title
    Write-Host "Established connection to SharePoint at $DeptUrl OK" -ForegroundColor Green
}
catch
{
    Write-Host "Not able to connect to SharePoint at $DeptUrl. Exception:$_.Exception.Message" -ForegroundColor red
    exit 1
}


#
