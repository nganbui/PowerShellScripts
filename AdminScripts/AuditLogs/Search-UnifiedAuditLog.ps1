$0 = $MyInvocation.MyCommand.Definition
$dp0 = [System.IO.Path]::GetDirectoryName($0)
$currentFolder = $dp0.Substring($dp0.LastIndexOf('\')+1)
$script:RootDir = Resolve-Path "$dp0\..\.." 
$logFileName = $MyInvocation.MyCommand.Name.Substring(0, $MyInvocation.MyCommand.Name.IndexOf('.'))
$logFileName = "$currentFolder\$logFileName"
$auditLogsFile = "AuditLogs.csv"
#Include core
."$script:RootDir\Common\Lib\LibCore.ps1"
."$script:RootDir\Common\Lib\LibO365.ps1"

Try {
    #log file path
    Set-LogFile -logFileName $logFileName
    $startTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    LogWrite -Message "----------------------- [Generate Audit Logs] Execution Started -----------------------"

    Set-TenantVars
    Set-AzureAppVars
    Set-DataFile    
    
    $StartDate = Read-host "Enter Start Date (m/d/yyyy)" 
    $EndDate = Read-host "Enter End Date (m/d/yyyy)" 

    $ReportGuest = [System.Collections.Generic.List[Object]]::new()
    Connect-ExchangeOnline
    $auditLogs = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -ResultSize 5000
    if ($auditLogs -ne $null) {
        $ConvertAudit = $auditLogs | Select-Object -ExpandProperty AuditData | ConvertFrom-Json        
    }
    LogWrite -Message  "Disconnect [EXO V2]"
    Disconnect-ExchangeOnline -Confirm:$false
    LogWrite -Message  "Export audit logs to .csv file"
    $ConvertAudit | Sort Name | Export-CSV -NoTypeInformation "$dp0\$auditLogsFile"
    
    
    $endTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    LogWrite -Message "[Populate Guest Users with their M365 Groups] Start Time: $($startTime)"
    LogWrite -Message "[Populate Guest Users with their M365 Groups] End Time:   $($endTime)"
    Write-Host "-------------------------------------------------------------------------"
    Write-host -ForegroundColor Green  "The file will be located at $dp0\$auditLogsFile"    
    
}
Catch [Exception] {
    LogWrite -Level ERROR "-Unexpected Error: $_ "
}
Finally {
    LogWrite -Message  "----------------------- [Generate Audit Logs] Completed ------------------------"
}
