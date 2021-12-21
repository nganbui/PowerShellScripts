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

Try {
    #-------- Set Global Variables ---------
    Set-TenantVars
    Set-AzureAppVars
    Set-DBVars    
    Set-LogFile -logFileName $logFileName
    Set-StatusVars
    Set-SiteRequestTypeVars
    Set-MiscVars
    Set-SNVars
            
    Update_SNIncident -IncidenttID "INC5562637" -IncidentType Provision -IncidentStatus Resolved -SiteURL "https://citspdev.sharepoint.com/sites/CIT-stdev"
    
}
Catch [Exception] {
    LogWrite -Level ERROR "[Unexpected Error]: $_ "
}
