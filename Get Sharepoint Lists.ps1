Import-Module 'C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll'
Import-Module 'C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll'
$site = 'https://dcids.sharepoint.com/sites/bi-site'
$admin = 'paul.davis@dcids.org'
$password = Read-Host 'Enter Passy pass pass' -AsSecureString



$context = New-Object Microsoft.SharePoint.Client.ClientContext($site)
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($admin, $password)
$context.Credentials = $credentials

try{
    $lists = $context.Web.Lists
    #$list = $lists.GetByTitle("Report Catalog")
    #$listItems = $list.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())

    $context.Load($lists)


    $context.ExecuteQuery()

    foreach($list in $lists){
        Write-Host $list.Title
        }    
}
catch{
        write-host "$($_.Exception.Message)"
}

