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
."$script:RootDir\Common\Lib\LibCache.ps1"
."$script:RootDir\Common\Lib\LibSPOSites.ps1"
."$script:RootDir\Common\Lib\LibSPOSitesDAO.ps1"
."$script:RootDir\Common\Lib\LibICs.ps1"
."$script:RootDir\Common\Lib\LibICsDAO.ps1"

<#
      .Synopsis
        Adding [ICName]_OD4BSecondaryAdmins and removing [OldICName]_OD4BSecondaryAdmins for Active Personal Sites
        Update SecondarySCA for PersonalSites table in DB
        - ICProfile DB Cache
        - PersonalSites DB Cache
        - Processing populate OD4B must get PersonalSites DB Cache because there is no IC info from O365 cache                     
    #>

Try {
    #log file path
    Set-LogFile -logFileName $logFileName
    Set-DataFile
    Set-DBVars
    Set-TenantVars
    Set-AzureAppVars
    $script:StartTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    LogWrite -Message "----------------------- [Populate Personal Site SCA] Execution Started -----------------------"
    #region Populate IC-OD4BSecondaryAdminsEmail for Personal Sites and update SecondarySCA for PersonalSites table in DB
    LogWrite -Message "-Getting IC Data..."
    $script:ICDBData = GetDataInCache -CacheType DB -ObjectType ICProfiles
    if($null -eq $script:ICDBData)
    {
        LogWrite -Message "ICs data not found in cache. Processing getting ICs from DB"
        SyncICProfileFromDBToCache
        $script:ICDBData = GetDataInCache -CacheType DB -ObjectType ICProfiles
    }
    LogWrite -Message "-Getting list of OD4BSecondaryAdminsEmail..."
    $OD4BSecondaryAdminsCol = $script:ICDBData | % { $_.OD4BSecondaryAdminsEmail.Trim() } | ? { $_ -ne '' }           
    $OD4BSecondaryAdminsCol = $OD4BSecondaryAdminsCol -split ' '

    LogWrite -Message "-Getting PersonalSites Data..."
    $script:PersonalSitesData = GetDataInCache -CacheType DB -ObjectType PersonalSites
    if($null -eq $script:PersonalSitesData)
    {
        LogWrite -Message "PersonalSites data not found in cache. Processing getting PersonalSites from DB"
        SyncPersonalSitesFromDBToCache
        $script:PersonalSitesData = GetDataInCache -CacheType DB -ObjectType PersonalSites
    }
    LogWrite -Message "-Processing populate Od4B secondary admins..."
    $script:PersonalSitesData = $script:PersonalSitesData | ? {$_.ICName -ne '' -and $_.ICName -ne 'Other'}    
    if($null -ne $script:PersonalSitesData) {        
        foreach($ps in $script:PersonalSitesData){
            $psURL = $ps.URL
            $OD4BSecondaryAdminsEmail = ($script:ICDBData | ? { $_.ICName -eq $ps.ICName}).OD4BSecondaryAdminsEmail.Trim()            
            ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $psURL                        
            $scaAdmins = (Get-PnPSiteCollectionAdmin | ? {$_.Email -ne ''}).Email            
            $scaAdmins = $scaAdmins -split ' '  
            LogWrite -Message " SCAs: $($scaAdmins)"
            # Remove oldIC_OD4BSecondaryAdminsEmails if find wrong IC_SCA 
            [System.Collections.ArrayList]$updatedSCA = @()
            $scaAdmins | % { 
                $sca = $_.Trim()
                $updatedSCA.Add($sca)                
                if ($OD4BSecondaryAdminsCol -contains $sca -and $OD4BSecondaryAdminsEmail -ne $sca) {                    
                    Remove-PnPSiteCollectionAdmin -Owners $sca
                    LogWrite -Message " Removed old IC_SecondaryAdmins [$($sca)] from the site [$($psURL)]" 
                    $updatedSCA.Remove($sca)            
                }
            } 
            # Adding IC_OD4BSecondaryAdminsEmails if not found from scaAdmins            
            if($scaAdmins -notcontains $OD4BSecondaryAdminsEmail) {               
               Add-PnPSiteCollectionAdmin -Owners $OD4BSecondaryAdminsEmail
               LogWrite -Message " Added $OD4BSecondaryAdminsEmail as SCA to [$($psURL)]" 
               $updatedSCA.Add($OD4BSecondaryAdminsEmail)
            }
            
            # Update SCA for PersonalSites table                
            $scaAdmins = $null                
            $scaAdmins = [string]::Join(";",$updatedSCA.ToArray())
            UpdatePersonalSiteSCA $script:connectionString $ps $scaAdmins
            LogWrite -Message " Updated [$($scaAdmins)] for Personal Site [$($psURL)]"
           
        }
    }    

    $script:EndTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    LogWrite -Message "[Populate Personal Site SCA] Start Time: $($script:StartTimeDailyCache)"
    LogWrite -Message "[Populate Personal Site SCA] End Time:   $($script:EndTimeDailyCache)"
    LogWrite -Message  "----------------------- [Populate Personal Site SCA] Execution Ended ------------------------"    
    #endregion
    
}
Catch [Exception] {
    LogWrite -Level ERROR "-Unexpected Error: $_ "
}
Finally {
    LogWrite -Message  "----------------------- [Populate Personal Site SCA] Completed ------------------------"
}
