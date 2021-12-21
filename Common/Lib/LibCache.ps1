Function GetDataInCache {
    param(
        [Parameter(Mandatory = $true)] [ValidateSet("O365", "DB")] $CacheType, 
        [Parameter(Mandatory = $true)] [ValidateSet("ICProfiles", "GraphAPIGroups", "GraphAPITeams", "GraphAPIChannel", "GraphAPIUsers", "SPOSites", "PersonalSites","PersonalSitesExtended", "O365Users", "O365Guests","Groups", "PortalAdminGroups", "SPOAdminGroups", "OneDriveAdminGroups","PowerBIworkspaces")] $ObjectType,
        [ValidateSet("Active", "InActive")] $ObjectState = "Active")
    $retData = @()
    try {
       
        $filePath = "$($script:CacheDataPath)\$($CacheType)\$($ObjectState)$($ObjectType).csv"
       

        if (test-path $filePath) {
            $retData = Import-csv $filePath            
        }
        else {
            LogWrite -Level INFO -Message "Cache File not found at $($filePath)"
        }        
    }
    catch {
        LogWrite -Level ERROR -Message "Error processing GetDataInCache(). Error info: $($_)"
        LogWrite -Level ERROR -Message "Error: $($_)"
    } 
    return $retData
}

Function SetDataInCache {
    param(
        [Parameter(Mandatory = $true)] $CacheData, 
        [Parameter(Mandatory = $true)] [ValidateSet("O365", "DB")] $CacheType, 
        [Parameter(Mandatory = $true)] [ValidateSet("ICProfiles", "GraphAPIGroups", "GraphAPITeams", "GraphAPIChannel", "GraphAPIUsers", "SPOSites", "PersonalSites","PersonalSitesExtended", "O365Users","O365Guests", "Groups", "PortalAdminGroups", "SPOAdminGroups", "OneDriveAdminGroups","PowerBIworkspaces")] $ObjectType,
        [ValidateSet("Active", "InActive", "ActiveNoAccess")] $ObjectState = "Active"
    )

    try {
        $cachePath = "$($script:CacheDataPath)\$($CacheType)"
        Create-Directory -dirPath $cachePath
        ExportCSV -DataSet $CacheData -FileName "$($cachePath)\$($ObjectState)$($ObjectType).csv"
        LogWrite -Level INFO -Message "$($CacheType) $($ObjectType) successfully cached"
    }
    catch {
        LogWrite -Level ERROR -Message "Error processing SetDataInCache(). Error info: $($_)"
    } 
}

<#
function GetExtDataInCache {
    param(
        [Parameter(Mandatory = $true)] [ValidateSet("SPOSites", "PersonalSites", "O365Users")] $ObjectType,
        [Parameter(Mandatory = $true)] $ICName, 
        [Parameter(Mandatory = $true)] $FileTimestamp, 
        [ValidateSet("Active", "InActive")] $ObjectState = "Active")
    
    try {
        $retData = $null

        $ConfigExtDataPath = "$($script:ConfigExtDataPath)\Retrieved"
        
        $filePath = "$($ConfigExtDataPath)\$($ObjectState)$($ObjectType)-$($ICName)-$($FileTimestamp).csv"

        #Write-log "$($filePath)"
        if (test-path $filePath) {
            $fileDate = (Get-ChildItem $filePath).LastWriteTime
            
            $retData = Import-csv $filePath
        }
        else {
            Write-Log "Cache File not found at $($filePath)"
        }
        
    }
    catch {
        Write-Log "Error processing GetDataInCache(). Error info: $($_)"
        Write-Host "Error: $($_)"
    } 
    return $retData
}

function SetExtDataInCache {
    param(
        [Parameter(Mandatory = $true)] $CacheData, 
        [Parameter(Mandatory = $true)] [ValidateSet("SPOSites", "PersonalSites", "O365Users")] $ObjectType,
        [Parameter(Mandatory = $true)] $ICName, 
        [Parameter(Mandatory = $true)] $FileTimestamp, 
        [ValidateSet("Active", "InActive")] $ObjectState = "Active"
    )

    try {
        $ConfigExtDataPath = "$($script:ConfigExtDataPath)\Retrieved"
        
        $filePath = "$($ConfigExtDataPath)\$($ObjectState)$($ObjectType)-$($ICName)-$($FileTimestamp).csv"
        
        if (test-path $ConfigExtDataPath) {
            Create_Directory -dirPath $ConfigExtDataPath
        }
        Export_CSV -DataSet $CacheData -FileName $filePath
        Write-Log "$($CacheType) $($ObjectType) Extended for $($ICName) successfully cached"
    }
    catch {
        Write-Log "Error processing SetDataInCache(). Error info: $($_)"
    } 
}

function SyncSPOSitesFromDBToCache {
    Write-Log "Syncing all DB Sites to Cache"
    #Get Data from DB
    $tempSites = GetSitesInDB -ConnectionString $script:ConnectionString -SitesType Sites -StatusType Active
    $tempDeletedSites = GetSitesInDB -ConnectionString $script:ConnectionString -SitesType Sites -StatusType InActive

    #Parse DB Data
    $script:activeSitesInDB = ParseSPOSites -sitesObj $tempSites -SitesType Sites -ObjectState Active -ParseObjType DB
    $script:deletedSitesInDB = ParseSPOSites -sitesObj $tempDeletedSites -SitesType Sites -ObjectState InActive -ParseObjType DB

    #Cache DB Sites Data
    SetDataInCache -CacheData $script:activeSitesInDB -CacheType DB -ObjectType SPOSites -ObjectState Active
    if ($script:deletedSitesInDB -ne $null) {
        SetDataInCache -CacheData $script:deletedSitesInDB -CacheType DB -ObjectType SPOSites -ObjectState InActive
    }
}

function SyncPersonalSitesFromDBToCache {
    Write-Log "Syncing all DB Personal Sites to Cache"

    #Get Data from DB
    $tempPersonalSites = GetSitesInDB -ConnectionString $script:ConnectionString -SitesType PersonalSites -StatusType Active
    $tempPersonalDeletedSites = GetSitesInDB -ConnectionString $script:ConnectionString -SitesType PersonalSites -StatusType InActive

    #Parse DB Data
    $script:activePersonalSitesInDB = ParseSPOSites -sitesObj $tempPersonalSites -SitesType PersonalSites -ParseObjType DB -ObjectState Active
    $script:deletedPersonalSitesInDB = ParseSPOSites -sitesObj $tempPersonalDeletedSites -SitesType PersonalSites -ObjectState InActive

    #Cache DB Sites Data
    SetDataInCache -CacheData $script:activePersonalSitesInDB -CacheType DB -ObjectType PersonalSites -ObjectState Active
    
    if ($script:deletedPersonalSitesInDB -ne $null) {
        SetDataInCache -CacheData $script:deletedPersonalSitesInDB -CacheType DB -ObjectType PersonalSites -ObjectState InActive
    }
}

function SyncO365UsersFromDBToCache {
    Write-Log "Syncing all DB Users to Cache"

    #Get Data from DB
    $tempUsers = GetUsersInDB -ConnectionString $script:ConnectionString -StatusType Active
    $tempDeletedUsers = GetUsersInDB -ConnectionString $script:ConnectionString -StatusType InActive

    Write-Log "Reterived the users from DB. Users count: $($tempUsers.count)"

    #Parse DB Data
    $script:activeUsersInDB = ParseO365Users -usersObj $tempUsers -ParseObjType DB
    $script:deletedUsersInDB = ParseO365Users -usersObj $tempDeletedUsers -ObjectState InActive -ParseObjType DB

    #Cache DB Sites Data
    SetDataInCache -CacheData $script:activeUsersInDB -CacheType DB -ObjectType O365Users -ObjectState Active

    if ($script:deletedUsersInDB -ne $null) {
        SetDataInCache -CacheData $script:deletedUsersInDB -CacheType DB -ObjectType O365Users -ObjectState InActive
    }
}

function SyncICProfileFromDBToCache {
    #Get Data from DB
    $tempICs = GetICsInDB -ConnectionString $script:ConnectionString

    #Parse DB Data
    $ICsInDB = ParseICs -ICsObj $tempICs

    #Cache DB Sites Data
    SetDataInCache -CacheData $ICsInDB -CacheType DB -ObjectType ICProfiles
}
#>
