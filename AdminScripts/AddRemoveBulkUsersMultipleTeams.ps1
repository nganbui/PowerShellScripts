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

try {
    #-------- Set Global Variables ---------
    Set-TenantVars
    Set-AzureAppVars
    Set-LogFile -logFileName $logFileName    
    #-------- Set Global Variables Ended --------- 
    #-------- Enter CSV location ---------
    $inputcsvFile = Import-Csv (Read-Host 'Enter CSV Location')
    #$csvFile = 'D:\Temp\BulkUsersTeams.csv'
    #$inputcsvFile = Import-Csv $csvFile
    Connect-MicrosoftTeams -TenantId $script:TenantId -ApplicationId  $script:appIdOperationSupport -CertificateThumbprint $script:appThumbprintOperationSupport

    foreach($line in $inputcsvFile){
        $t = $line.GroupId                
        $OwnerEmailToAdd = $line.OwnerEmailToAdd.Split(";")        
        $OwnerEmailToRemove = $line.OwnerEmailToRemove.Split(";")        
        <#--Add owners from Teams--#>        
        if ($OwnerEmailToAdd -and $OwnerEmailToAdd.Length -gt 1){
            LogWrite -Level INFO -Message "Adding owners to the team [$($t)]..."
            foreach($userAdded in $OwnerEmailToAdd){
                LogWrite -Level INFO -Message "Adding user $($userAdded)"
                Add-TeamUser -GroupId $t -User $userAdded -Role Owner -ErrorAction SilentlyContinue            
            }
        }
        <#--Remove owners from Teams--#>   
        if ($OwnerEmailToRemove -and $OwnerEmailToRemove.Length -gt 1){
            LogWrite -Level INFO -Message "Removing owners from [$($t)]..."
            foreach($userRemoved in $OwnerEmailToRemove){
                LogWrite -Level INFO -Message "Removing user $($userRemoved)"
                Remove-TeamUser -GroupId $t -User $userRemoved -ErrorAction SilentlyContinue            
            }
        }
    } 
}  
catch {
    LogWrite -Level ERROR "Error in the script: $($_)"
}
finally{
    LogWrite -Level INFO -Message "Disconnect Microsoft Teams."
    DisconnectMicrosoftTeams   
}