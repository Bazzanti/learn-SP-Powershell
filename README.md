# learn-Sp-Powershell
Exercises and scripts about Sharepoint Powershell

SHAREPOINT SHELL AS ADMINISTRATOR
SCRIPT PER MODIFICARE IL VALORE DI UN CAMPO SU n ITEMS IN UNA LISTA SHAREPOINT -

$URL = "<URL HERE>"
$ListName = "<listname>"
$Web = Get-SPWeb $URL
$user = $Web.EnsureUser("<USER HERE>")

$List = $Web.Lists.TryGetList($ListName)

$query = New-Object Microsoft.Sharepoint.SPQuery
$query.Query = "
<Where>
  <And>
    <Eq>
      <FieldRef Name='Tutor' LookupId='True'></FieldRef>
      <Value Type='Lookup'>117</Value>
    </Eq>
    <Neq>
      <FieldRef Name='Tutor' LookupId='True'></FieldRef>
      <Value Type='Lookup'>33</Value>
    </Neq>
  </And>
</Where>"
$query.RowLimit = 1000
$items = $List.GetItems($query)
Write-Host "Found $($items.count) elements"

#Controllo
for($i=$items.Count-1;$i -ge 0;$i--){
  $item = $items[$i]
    
   Write-Host $i - $item.ID - $item["Tutor"]
  
}

#Update
for($i=$items.Count-1;$i -ge 0;$i--){
  $item = $items[$i]
    
    $itemTU = $List.Items.GetItemById($item.ID);
    $itemTU["Tutor"] = $user ;
    $itemTU.Update()
    Write-Host $i - $item.ID - Old: $item["Tutor"] - New: $itemTU["Tutor"] 
}


____________________________________________________________________

# GET ITEM BY ID
$ListItem = $list.GetItemById(ID)

#GET FIELD FROM ITEM
  $Title = $ListItem["Title"]

  #GET DATE FIELD
  Get-Date ($ListItem["Modified"]) -Format "dd-MMM-yyyy"
  
  #GET LOOKUP FIELD
  $Lookup = New-Object Microsoft.SharePoint.SPFieldLookupValue($Item[$LookupFieldName])
  write-host $Lookup.LookupValue

  #GET USER FIELD
  $CreatedBy = $ListItem["Created By"]
  $CreatedByUserObj = New-Object Microsoft.SharePoint.SPFieldUserValue($web, $CreatedBy)
  $CreatedByDisplayName = $CreatedByUserObj.User.DisplayName;

#SET FIELD 
  $ListItem["Field"] = "Value"
  $ListItem.Update()

  #SET DATE
  $StringNewStartTime = 11/25/2013 11:00 PM
  #SET LOOKUP FIELD 
  $ParentListName="Parent Projects" #Lookup Parent List
  $ChildListName="Project Milestones" #List to add new lookup value
  $ParentListLookupField="Project Name"
  $ChildListLookupField ="Parent Project"
  $ParentListLookupValue="Cloud Development" #Parent Project value

  $ParentListLookupItem = $ParentList.Items | where {$_[$ParentListLookupField] -eq $ParentListLookupValue}
  $NewItem[$ChildListLookupField] = $ParentListLookupItem.ID
  $NewItem.Update()   

  #SET USER FIELD

  $UserAccount="USERACCOUNT"
  #To Add new List Item, use $Item = $List.AddItem()
  $User = Get-SPUser -Identity $UserAccount -Web $web
  $ListItem[$FieldName] = $User
  $ListItem.Update()

#OPERAZIONE NESTATA CHE STAMPA 100 ELEMENTI PER VOLTA

  ## View XML
  $qCommand = @"
  <View Scope="RecursiveAll">
      <Query>
          <OrderBy><FieldRef Name='ID' Ascending='TRUE'/></OrderBy>
      </Query>
      <RowLimit Paged="TRUE">100</RowLimit>
  </View>
  "@

  $camlQuery = New-Object Microsoft.Sharepoint.SPQuery
  $camlQuery.ListItemCollectionPosition = $position
  $camlQuery.ViewXml = $qCommand

  ##items
  Do{    
    $currentCollection = $List.GetItems($camlQuery)
    $camlQuery.ListItemCollectionPosition = $currentCollection.ListItemCollectionPosition
    
    Write-Host
    Write-Host "Start ciclo" $camlQuery.ListItemCollectionPosition.PagingInfo

    #Controllo
    for($i=$currentCollection.Count-1;$i -ge 0;$i--){
       $item = $currentCollection[$i]
       Write-Host $i - $item.ID - $item["Title"]
    }
  }
  # the position of the last page will be Null
  Until($camlQuery.ListItemCollectionPosition -eq $null ) 


#PRINT CSV

#Array to Hold Result - PSObjects
  $ListItemCollection = @()

   #Get All List items where ContactManagerArea is "East"
   $List.Items | Where-Object { $_["ContactManagerArea"] -eq "East"} |  foreach { 
   
    #Get user value
    $CreatedBy = $_["Created By"]
    $CreatedByUserObj = New-Object Microsoft.SharePoint.SPFieldUserValue($web, $CreatedBy)
    $CreatedByDisplayName = $CreatedByUserObj.User.DisplayName;

    #Get Lookup value
    $ContactIntercompany = New-Object Microsoft.SharePoint.SPFieldLookupValue($_["ContactIntercompany"])

   $ExportItem = New-Object PSObject
   $ExportItem | Add-Member -MemberType NoteProperty -name "Title" -value $_["Title"]
   $ExportItem | Add-Member -MemberType NoteProperty -name "ManagerArea" -value $_["ContactManagerArea"]
   $ExportItem | Add-Member -MemberType NoteProperty -Name "Modified" -value $_["Modified"]
   $ExportItem | Add-Member -MemberType NoteProperty -Name "Created By" -value $CreatedByDisplayName
   $ExportItem | Add-Member -MemberType NoteProperty -name "Intercompany" -value $ContactIntercompany.LookupValue

   #Add the object with property to an Array
   $ListItemCollection += $ExportItem
   }
   #Export the result Array to CSV file
   $ListItemCollection | Export-CSV "C:\Users\<user>\Desktop\ListData.csv" -NoTypeInformation                       
   
  #Dispose the web Object
  $web.Dispose()


#QUERY 

$query = New-Object Microsoft.Sharepoint.SPQuery
$query.Query = "<Query>
  <Where>
  </Where>
</Query>"



  #SIMPLE QUERY

  <Query>
  <Where>
    <Eq>
      <FieldRef Name="Title"></FieldRef>
      <Value Type="Text">Test</Value>
    </Eq>
  </Where>
  </Query>

  #QUERY BETWEEN DATES
      <And>
         <Geq>
            <FieldRef Name='Modified' />
            <Value Type='DateTime'>2014-08-01T12:00:00Z</Value>
         </Geq>
         <Leq>
            <FieldRef Name='Modified' />
            <Value Type='DateTime'>2019-08-01T12:00:00Z</Value>
         </Leq>
      </And>

  #QUERY FOR USER
    <Eq>
      <FieldRef Name='Tutor' LookupId='True'></FieldRef>
      <Value Type='Lookup'>117</Value>
    </Eq>


  #  HOW TO NEST
    <Where>
          <And>       
              <Or>
                  <Eq>
                      <FieldRef Name='FirstName' />
                      <Value Type='Text'>Doe</Value>
                  </Eq>
                  <Or>
                      <Eq>
                          <FieldRef Name='LastName' />
                          <Value Type='Text'>Doe</Value>
                      </Eq>
                      <Eq>
                          <FieldRef Name='Profile' />
                          <Value Type='Text'>Doe</Value>
                      </Eq>
                  </Or>
              </Or>
              <Or>
                  <Eq>
                      <FieldRef Name='FirstName' />
                      <Value Type='Text'>123</Value>
                  </Eq>
                  <Or>
                      <Eq>
                          <FieldRef Name='LastName' />
                          <Value Type='Text'>123</Value>
                      </Eq>
                      <Eq>
                          <FieldRef Name='Profile' />
                          <Value Type='Text'>123</Value>
                      </Eq>
                  </Or>
              </Or>
          </And>
    </Where>


___________________________


# SCRIPT


$user = $w.EnsureUser("<user>")

$URL = "<url>"
$ListName = "<listname>"
$Web = Get-SPWeb $URL
$List = $Web.Lists.TryGetList($ListName)
$LocationTypeList = $Web.Lists.tryGetList("LocationTypes") 
$AirportItem = $LocationTypeList.Items | where {$_["Title"] -eq "Airport"}
$AirportLookup = New-Object Microsoft.Sharepoint.SPFieldLookupValue($AirportItem.ID, $AirportItem["Title"])
$query = New-Object Microsoft.Sharepoint.SPQuery
$query.Query = "<Where></Where>"
$query.RowLimit = 100000
$items = $List.GetItems($query)
Write-Host "Found $($items.count) elements"
#Controllo
for($i=$items.Count-1;$i -ge 0;$i--){
  $item = $items[$i]
    
  $airport = $item["Airport"];
  if ($airport -eq $true) {
    Write-Host $i - $item.ID - $airport - $item["BuildingLocationType"]
  }
}
#Update
for($i=$items.Count-1;$i -ge 0;$i--){
  $item = $items[$i]
    
  $airport = $item["Airport"];
  if ($airport -eq $true) {
    $itemTU = $List.Items.GetItemById($item.ID);
    $itemTU["BuildingLocationType"] = $AirportLookup;
    $itemTU.Update()
    Write-Host $i - $item.ID - $airport - $item["BuildingLocationType"]
  }
}
