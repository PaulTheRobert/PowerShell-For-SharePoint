<#
    created by: paul davis
    created on: 04/27/21

    purpose: this script is the first pass at automating the load of a Sharepoint List: Report Catalog from a SQL qry against the DEV and PROD BI servers.

    yeet . . .

    first pass only looking to create and maintain the basic catalog, not doing any update if record already exists


    1. Qry PROD& DEV
    2. Check / Load PROD reports into Sharepoint
    3. Check / Load DEV reports into Sharepoint joined on Path (<i know its wierd but kinda makes sense)
    
#>

######################################################################################################################################
#Get Report Catalog from DEV

$PRODConnectionString = "Data Source=DCIDS-BI-PROD1;Integrated Security=True;ApplicationIntent=ReadOnly"
$DEVConnectionString = "Data Source=DCIDS-BI-DEV1;Integrated Security=True;ApplicationIntent=ReadOnly"

$Qry = "USE [PowerBI] SELECT CAT.[ItemId], CAT.[Name] AS [Title], CAT.[Path], CASE CAT.[Type] WHEN 2 THEN 'SSRS' WHEN 13 THEN 'PowerBI' END AS [Type] FROM	[dbo].[Catalog] CAT WHERE CAT.[Type] IN( 2, 13) ORDER BY CAT.[Name]"

$DevReports = Invoke-Sqlcmd -Query $Qry -ConnectionString $DEVConnectionString 
$ProdReports = Invoke-Sqlcmd -Query $Qry -ConnectionString $PRODConnectionString 

######################################################################################################################################
# these dlls are for interacting with sharepoint 

Import-Module 'C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll'
Import-Module 'C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll'

$site = 'https://dcids.sharepoint.com/sites/bi-site'
$admin = Read-Host 'Enter Username (user.name@dcids.org) '
$password =  Read-Host 'Enter Password '   -AsSecureString

$context = New-Object Microsoft.SharePoint.Client.ClientContext($site)
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($admin, $password)
$context.Credentials = $credentials

######################################################################################################################################

#declare a collection of psobjects for sharepoint items
$sharePointListItems = @()

#get the sarting sharepoint list into an ps object
try{
    $lists = $context.Web.Lists
    $list = $lists.GetByTitle("Report Catalog")

    $listItems = $list.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())
    $context.Load($listItems)

    $context.ExecuteQuery()
    
    foreach($item in $listitems){
        #create PS Object for later processing help
        $sharePointListItem = [PSCustomObject]@{                
            'ID'              = $item["ID"] 
            'PROD_ItemId'       = $item["PROD_ItemId"]
            'DEV_ItemId'        = $item["DEV_ItemId"]
            'Title'             = $item["Title"]   
            'Path'              = $item["Path"]         
            }
                
        #add the sharepoint item psobject to the collection
        $sharePointListItems += $sharePointListItem 
       }
    }
    catch{
            write-host "$($_.Exception.Message)"
}

#loop through the PROD result set
foreach($Report in $ProdReports){
    $exists = 0

    #does this report already exist in the catalog?
    foreach($item in $sharePointListItems){
        #write-host $Report.ItemId " | " $item.PROD_ItemId " | " $Report.Title
        if($Report.ItemId -eq $item.PROD_ItemId){
            $exists = 1            
        }        
    }

    #if the report does not already exist in the catalog, add it
    if($exists -eq 0){
    write-host "Writing " $Report.Title " to the the PROD catalog"
        try{
            $lists = $context.Web.Lists
            $list = $lists.GetByTitle("Report Catalog")
            $listItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation      
            
            $listItem = $list.AddItem($listItemInfo)
            $listItem["PROD_ItemId"] = $Report.ItemId
            $listItem["Title"] = $Report.Title
            $listItem["Path"] = $Report.Path
            $listItem["ReportType"] = $Report.Type
            $listItem.Update()

            $context.Load($list)
            $context.ExecuteQuery()
             
        }
        catch{
                write-host "$($_.Exception.Message)"
        }
    } 
    #if the report does exist in the catalog leave it alone
    else{
        write-host $Report.Title " is already in the catalog"
    }
}

######################################################################################################################################
#loop through the DEV result set

foreach($Report in $DevReports){
    #initialize flag for if already exists
    $exists = 0

    #does this report already exist in the catalog?

    # Match 1st by Path
    # Match 2nd by DEV_ItemId
     #The assumption is that DEV and PROD will have identical paths most of the time, DEV and PROD dont always sync up perfectly.

    foreach($item in $sharePointListItems){
        # Match 1st by Path
        if($Report.Path -eq $item.Path){
            $exists = 1
        
            # Match 2nd by DEV_ItemId 
            # should his elseif instead be nested inside of the above if?  
            if($Report.ItemId -eq $item.DEV_ItemId){
                $exists = 2
            }
        }
    }

    Switch($exists){
        # add the DEV report to the Catalog
        0 {
            write-host "Writing " $Report.Title " to the the DEV catalog"
            try{
                $lists = $context.Web.Lists
                $list = $lists.GetByTitle("Report Catalog")
                $listItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation      
            
                $listItem = $list.AddItem($listItemInfo)
                $listItem["DEV_ItemId"] = $Report.ItemId
                $listItem["Title"] = $Report.Title
                $listItem["Path"] = $Report.Path
                $listItem["ReportType"] = $Report.Type
                $listItem.Update()

                $context.Load($list)
                $context.ExecuteQuery()             
            }
            catch{
                write-host "$($_.Exception.Message)"
            }        
        }

        #Update the DEV_ItemId on the Prod Report Catalog Item
        1 {
            Write-Host $Report.Title " Updtating DEV_ItemId" 
            try{
                $lists = $context.Web.Lists
                $list = $lists.GetByTitle("Report Catalog")
                $listItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation

                $listItem = $list.GetItemById($item.ID)        
                $listItem["DEV_ItemId"] = $Report.ItemId
                $listItem.Update()

                $context.Load($list)
                $context.ExecuteQuery()  
            }
            catch{
                    write-host "$($_.Exception.Message)"
            }
        
        }

        #Do nothing - Dev report already existis in Reeport Catalog
        2 {
            Write-Host $Report.Title " already exists in the report catalog and is up to date." 
            break
           }


    }
}
    
