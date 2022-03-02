Import-Module 'C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll'
Import-Module 'C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll'
$site = 'https://dcids.sharepoint.com/sites/bi-site'
$admin = 'paul.davis@dcids.org'
$password =  Read-Host 'Enter Passy pass pass'   -AsSecureString

$context = New-Object Microsoft.SharePoint.Client.ClientContext($site)
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($admin, $password)
$context.Credentials = $credentials

try{
    $lists = $context.Web.Lists
    $list = $lists.GetByTitle("Report Catalog")

    #$listItems = $list.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())
    #$context.Load($listItems)

    $context.Load($list.Fields)

    $context.ExecuteQuery()

##    foreach($Item in $listItems){
##        Write-Host "Title:" $Item["Title"] "DevItemId:" $Item["DevItemId"] "DevPath:" $Item["DevPath"] "DevItemId:" $Item["ReportType"]
##        }   

    foreach($field in $list.Fields){
        Write-Host "Title: " $field.Title " | Internal Name: " $field.InternalName " | indexed? " $field.Indexed
    }
 
}
catch{
        write-host "$($_.Exception.Message)"
}

