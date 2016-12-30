#
# window.ps1
Add-PSSnapin WASP
$url = "https://kabali11.sharepoint.com/sites/test1/" 
$username="admin@kabali11.onmicrosoft.com" 
$password="Infy@123" 


$ie = New-Object -com internetexplorer.application; 
$ie.visible = $true; 
$ie.navigate($url);

Import-Module WASP
while ($ie.Busy -eq $true) 
{ 
    Start-Sleep -Milliseconds 1000; 
} 

#$ie.Document.getElementById("cred_userid_inputtext").value = $username 
#$ie | Send-Keys "{TAB}"
#Start-Sleep -Milliseconds 4000

$Link1=$ie.Document.getElementByID("admin_kabali11_onmicrosoft_com") 
$Link1.click()
Start-Sleep -Milliseconds 4000
$ie.Document.getElementByID("cred_password_inputtext").value=$password 

$Link=$ie.Document.getElementByID("cred_sign_in_button") 
$Link.click()
Start-Sleep -Milliseconds 10000

$ie.Navigate("https://kabali11.sharepoint.com/sites/test1/_layouts/15/AreaNavigationSettings.aspx")

Start-Sleep -Milliseconds 10000

$Link2=$ie.Document.getElementByID("edit_link") 
$Link2.click()

Start-Sleep -Milliseconds 4000




#
