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
    #$inputcsvFile = Import-Csv (Read-Host 'Enter CSV Location')
    $csvFile = 'D:\Temp\BulkUsersTeams.csv'
    $inputcsvFile = Import-Csv $csvFile
    $Certificate = Get-Item Cert:\LocalMachine\My\* | Where-Object { $_.Subject -ieq "CN=SPO-M365OperationSupport.nih.sharepoint.com" }
    $token = Connect-NIHO365GraphWithCert -TenantName $script:TenantId -AppId $script:appIdOperationSupport -Certificate $Certificate        

    foreach($line in $inputcsvFile){
        $g = $line.GroupId                
        $UPNsAdded = $line.UPNsAdded -split ';'        
        $UPNsAdded = @($UPNsAdded)
        $UPNsRemoved = $line.UPNsRemoved -split ';'
        $UPNsRemoved = @($UPNsRemoved)
        $role = 'Onwer'
        if ($line.Role -ne $null -or $line.Role -ne ''){
            $role = $line.Role
        }
        <#--Add users from Groups--#>        
        if ($UPNsAdded -and $UPNsAdded.Count -gt 0){
            LogWrite -Level INFO -Message "Adding users to the group [$($g)]..."
            foreach($userAdded in $UPNsAdded){
                LogWrite -Level INFO -Message "Adding user $($userAdded)"
                Add-NIHO365GroupMember -AuthToken $token -Group $g -Member $userAdded -AsOwner
            }
        }
        <#--Remove users to Groups--#>        
        if ($UPNsRemoved -and $UPNsRemoved.Count -gt 0){
            LogWrite -Level INFO -Message "Removing users from the group [$($g)]..."
            foreach($userRemoved in $UPNsRemoved){
                LogWrite -Level INFO -Message "Removing user $($userRemoved)"
                Remove-NIHO365GroupMember -AuthToken $token -Group $g -Member $userRemoved -AsOwner         
            }
        }
    } 
}  
catch {
    LogWrite -Level ERROR "Error in the script: $($_)"
}
