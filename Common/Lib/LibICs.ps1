Function ParseICs {
    param($ICsObj)
    
    $ICsFormattedData = @()

    foreach ($IC in $ICsObj) {        
        
        $ICsFormattedData += [pscustomobject]@{
            ICName                     = $IC.ICName;
            Aliases                    = $IC.Aliases;
            O365PortalAdminsEnabled    = $IC.O365PortalAdminsEnabled;
            O365PortalAdmins           = $IC.O365PortalAdmins;
            SPOSecondaryAdminsEnabled  = $IC.SPOSecondaryAdminsEnabled;
            SPOSecondaryAdmins         = $IC.SPOSecondaryAdmins;
            SPOSecondaryAdminsEmail    = $IC.SPOSecondaryAdminsEmail;
            OD4BSecondaryAdminsEnabled = $IC.OD4BSecondaryAdminsEnabled;
            OD4BSecondaryAdmins        = $IC.OD4BSecondaryAdmins;
            OD4BSecondaryAdminsEmail   = $IC.OD4BSecondaryAdminsEmail;
            Comments                   = $IC.Comments;        
            Operation                  = "";
            OperationStatus            = "";
            AdditionalInfo             = ""
        }
    }
    return $ICsFormattedData
}

Function SyncICProfileFromDBToCache{
    #Get Data from DB
    $ICsInDB = GetICsInDB -ConnectionString $script:ConnectionString

    #Parse DB Data
    $ICsInDB = ParseICs -ICsObj $ICsInDB

    #Cache DB Sites Data
    if ($null -ne $ICsInDB) {
        SetDataInCache -CacheData $ICsInDB -CacheType DB -ObjectType ICProfiles
    }
}
