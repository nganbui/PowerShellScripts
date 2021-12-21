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
      ===========================================================================
      .DESCRIPTION
        Adding [ICName]_SPOSecAdmins as SCA for SPO Sites
        Update SecondarySCA for SPO Sites in DB
        - ICProfile DB Cache
        - Sites DB Cache
        - Processing populate [ICName]_SPOSecAdmins must get Sites DB Cache because there is no IC info from O365 cache                     
    #>

Try {
    #log file path
    Set-LogFile -logFileName $logFileName
    Set-DataFile
    Set-DBVars
    Set-TenantVars
    Set-AzureAppVars

    $script:StartTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    LogWrite -Message "----------------------- [Process SPO Sites] Execution Started -----------------------"
       

    $script:EndTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    LogWrite -Message "[Process SPO Sites] Start Time: $($script:StartTimeDailyCache)"
    LogWrite -Message "[Process SPO Sites] End Time:   $($script:EndTimeDailyCache)"
    LogWrite -Message  "----------------------- [Populate Personal Site SCA] Execution Ended ------------------------"    
   
    
}
Catch [Exception] {
    LogWrite -Level ERROR "-Unexpected Error: $_ "
}
Finally {
    LogWrite -Message  "----------------------- [Process SPO Sites] Completed ------------------------"
}
