$0 = $MyInvocation.MyCommand.Definition
$dp0 = [System.IO.Path]::GetDirectoryName($0)
$currentFolder = $dp0.Substring($dp0.LastIndexOf('\')+1)
$script:RootDir = Resolve-Path "$dp0\.." 
$logFileName = $MyInvocation.MyCommand.Name.Substring(0, $MyInvocation.MyCommand.Name.IndexOf('.'))
$logFileName = "$currentFolder\$logFileName"
#Include core
."$script:RootDir\Common\Lib\LibCore.ps1"
#Include dependent functionality
."$script:RootDir\Common\Lib\LibO365.ps1"
."$script:RootDir\Common\Lib\ProvisionHelper.ps1"
."$script:RootDir\Common\Lib\GraphAPILibO365Groups.ps1"
."$script:RootDir\Common\Lib\GraphAPILibO365GroupsDAO.ps1"
."$script:RootDir\Common\Lib\GraphAPILibSPOSites.ps1"
."$script:RootDir\Common\Lib\GraphAPILibSPOSitesDAO.ps1"
."$script:RootDir\Common\Lib\GraphAPILibO365UsersDAO.ps1"
."$script:RootDir\Common\Lib\LibRequestDAO.ps1"
."$script:RootDir\Common\Lib\LibUtils.ps1"
."$script:RootDir\Common\Lib\LibPowerBIWorkspace.ps1"
."$script:RootDir\Common\Lib\LibPowerBIWorkspaceDAO.ps1"


Function Process-Provision {
    param([Parameter(Mandatory = $true)] $Requests)
    
    #$templates = @('Team','SPO') # or get from SiteTemplates table
    # First run: provision Team/SPO with basic settings and assign svc as an owner + update request status to "In Progress"
    # Second run: add business owner as group/site owner;update storage quota and description;sync to DB;update request status to "Completed" and SN ticket plus send email confirmation to requestor/owner
    # https://developer.microsoft.com/en-us/outlook/blogs/announcing-support-for-new-groups-properties-via-microsoft-graph-api/

    foreach ($req in $Requests) {
        #Get Signin Name by email 
        $reqOwner = $req.PrimarySCA.Trim()
        $reqStatus = $req.RequestStatusId.Guid
        $reqSiteURL = $req.SiteUrl
        if ($null -eq $reqSiteURL -or [string]::Empty -eq $reqSiteURL){
            $reqSiteURL = $req.SiteName
        }
        $SiteOwnerId = @()
        $SiteOwnerUPN = @()
        # Primary Owner
        if (![string]::IsNullOrWhiteSpace($reqOwner)){
           $dt = GetSigninNameByEmail -Email $reqOwner -connectionString $script:ConnectionString
           $SiteOwnerId+= $dt["UserId"].Guid
           $SiteOwnerUPN+= $dt["SigninName"].Trim()

        }               
        # Secondary owner
        if (![string]::IsNullOrWhiteSpace($req.SecondaryOwnerEmail)){
            $dtSecondOwner = GetSigninNameByEmail -Email $req.SecondaryOwnerEmail -connectionString $script:ConnectionString
            if ($dtSecondOwner["UserId"].Guid -ne $SiteOwnerId){
                $SiteOwnerId+= $dtSecondOwner["UserId"].Guid
                $SiteOwnerUPN+= $dtSecondOwner["SigninName"].Trim()
            }
        }        

        $siteOwners = @{           
            SiteOwnerId = $SiteOwnerId
            SiteOwnerUPN = $SiteOwnerUPN
        }
        $request = ParseRequest $req $siteOwners
        
        try{
            switch ($reqStatus) {
                $script:Submitted{                
                    LogWrite " $($script:ProcessNew): New Request: [$reqSiteURL]"
                    Provision-New -Request $request
                }
                $script:InProgress{
                    LogWrite " $($script:ProcessInProgress): Pending Request: [$reqSiteURL]"                
                    Provision-InProgress -Request $request
                }
            }
            LogWrite " #----------------------------------------------------------------------#"
        }
        catch{
            LogWrite -Level ERROR "[Unexpected Error]:: Exit the loop if encountered an unexpected error."
            break
        }
    }
}

Function Process-Decommission {
    param([Parameter(Mandatory = $true)] $Requests)

    foreach ($req in $Requests) {        
        $reqStatus = $req.RequestStatusId.Guid
        $reqSiteURL = $req.SiteUrl

        try{
            switch ($reqStatus) {
                $script:Submitted{                
                    LogWrite " $($script:ProcessDecommission): Decommission Site Request: [$reqSiteURL]"
                    Decommission-SPOSite -Request $req
                    
                }                
            }
            LogWrite " #----------------------------------------------------------------------#"
        }
        catch{
            LogWrite -Level ERROR "[Unexpected Error]:: Exit the loop if encountered an unexpected error."
            break
        }    
    }
}

Try {
    #-------- Set Global Variables ---------
    Set-TenantVars
    Set-AzureAppVars
    Set-DBVars    
    Set-LogFile -logFileName $logFileName        
    #-------- Set Global Variables Ended ---------    
    
    $startTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    LogWrite -Message "----------------------- [Processing site requests] Started -----------------------"    
    #-All active requests including provision and decomission-#
    $activeRequests = @(GetActiveSiteRequests $script:ConnectionString)
    if ($activeRequests.Count -eq 0) {       
        LogWrite -Message "No new requests found."        
    }
    else {        
        #-- Initialize script scope variables --#
        Set-StatusVars
        Set-SiteRequestTypeVars
        Set-MiscVars
        Set-SNVars
        Set-EmailVars
        #--===============================--#

        $pendingRequests = @($activeRequests | ? { $_.ProcessFlag -eq 1 })
        $activeRequests = @($activeRequests | ? { $_.ProcessFlag -ne 1 } | select -first $($script:MaxRequests))

        $provisionRequests = @($activeRequests  | ? { $_.RequestTypeId -eq $script:Provision })
        $decomissionRequests = @($activeRequests  | ? { $_.RequestTypeId -eq $script:Decomission })
        #-Call provision script
        if ($provisionRequests.Count -gt 0){
            Process-Provision -Requests $provisionRequests
        }
        #-Call decommission script
        if ($decomissionRequests.Count -gt 0){
            Process-Decommission -Requests $decomissionRequests
        }
    }

    $endTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"    
    LogWrite -Message " [Teams/SPO Site Provisioning] Start Time: $startTime"
    LogWrite -Message " [Teams/SPO Site Provisioning] End Time:   $endTime"
    LogWrite -Message  "----------------------- [Processing site requests] Ended ------------------------"  
    
}
Catch [Exception] {
    LogWrite -Level ERROR "[Unexpected Error]: $_ "
}
Finally {
    if($script:TenantContext){
        LogWrite -Level INFO -Message " Disconnect SharePoint Admin Center." 
        DisconnectPnpOnlineOAuth -Context $script:TenantContext 
    }    
        
}