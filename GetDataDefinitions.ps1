<#
    created by: paul davis
    created on: 09/09/2021

    - creates and populates data from SharePoint Data Def. Cat. --> [ExternalSources].[SharePoint].[DataDefinitionCatalog]

    - This script is designed to pull the data from the SharePoint data definition catalog, and transform it to be easily accesibe for reporting purposes.
    - The script will drop the whole table, create it, then populate it each time. This makes it easy for a future developer to modify the table from within this script.
   
#>
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
    $list = $lists.GetByTitle("Data Definition Catalog")

    $listItems = $list.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())
    $context.Load($listItems)

    $context.ExecuteQuery()

    ##declare an empty collection to hold the DataElement objects
    $ReportDataDictionary = New-Object System.Collections.ArrayList

    foreach($item in $listitems){

     ## loop through each report that is tied to this data def
     foreach($Report in $item["Report"]){
        
        ##declare an empty custom ps object to store the data dictionary for each report post tranformation. This will serve as staging for the sql load    
        $DataElement = @{
            reportId          = $Report.LookupId
            reportName        = $Report.LookupValue
            DataElementId     = $item["ID"]
            DataElementTitle  = $item["Title"]
            DataElementType   = $item["DataElementType"]
            DataElementStatus = $item["DataElementStatus"]
            Description       = $item["Description"]
            BusinessLogic     = $item["BusinessLogic"]
            Grain             = $item["Grain"]     
            }

        $ReportDataDictionary.Add($DataElement)
                   
        }          
     
  }

  #this is the connection string to dev
  $ConnectionString = "Data Source=DCIDS-SQL-DEV1;Integrated Security=True;ApplicationIntent=ReadOnly"

  #drop table
  $Qry = "DROP TABLE [ExternalSources].[SharePoint].[DataDefinitionCatalog];"
  $SqlResult = Invoke-Sqlcmd -Query $Qry -ConnectionString $ConnectionString

  #create table
  $Qry = "CREATE TABLE [ExternalSources].[SharePoint].[DataDefinitionCatalog] (reportId nvarchar(max), reportName nvarchar(max), DataElementId nvarchar(max), DataElementTitle nvarchar(max), DataElementType nvarchar(max), DataElementStatus nvarchar(max), Description nvarchar(max), BusinessLogic nvarchar(max), Grain nvarchar(max));"
  $SqlResult = Invoke-Sqlcmd -Query $Qry -ConnectionString $ConnectionString
  
  ##dynamically build and execute sql statments to insert record
  ForEach($item in $ReportDataDictionary){
    $Qry = "USE [ExternalSources] INSERT INTO [ExternalSources].[SharePoint].[DataDefinitionCatalog] VALUES($($item.reportId), '$($item.reportName)', $($item.DataElementId), '$($item.DataElementTitle)', '$($item.DataElementType)', '$($item.DataElementStatus)', '$($item.Description)', '$($item.BusinessLogic)', '$($item.Grain)')"
    write-host $qry
    $SqlResult = Invoke-Sqlcmd -Query $Qry -ConnectionString $ConnectionString 
  }
  
  
}
catch{
        write-host "$($_.Exception.Message)"
}

