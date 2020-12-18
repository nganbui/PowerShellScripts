$0 = $MyInvocation.MyCommand.Definition
$dp0 = [System.IO.Path]::GetDirectoryName($0)
$currentFolder = $dp0.Substring($dp0.LastIndexOf('\')+1)
$script:RootDir = Resolve-Path "$dp0\.." 
$logFileName = $MyInvocation.MyCommand.Name.Substring(0, $MyInvocation.MyCommand.Name.IndexOf('.'))
$logFileName = "$currentFolder\$logFileName"
$usageReport = "UsageReports"
$inputReport = "Input"
#Include core
."$script:RootDir\Common\Lib\LibCore.ps1"
#Include dependent functionality
."$script:RootDir\Common\Lib\LibO365.ps1"
."$script:RootDir\Common\Lib\LibCache.ps1"

<# 
    The report should run on the second day of a month
    Reports are available for the last 7 days, 30 days, 90 days, and 180 days. 
    Data won't exist for all reporting periods right away. 
    The reports become available within 48 hours.
#>

Function GetAllM365Reports {
    #----------------Read report endpoint-----------------#            
    $path = "$dp0\ReportsMetadata.psd1"
    $reportMetaData = Import-PowerShellDataFile -Path $path
    #----------------End-Read report endpoint-------------#
    
    #--Create a folder UsageReports under Data if any        
    $date = Get-Date
    $year = $date.Year
    $month = $date.AddMonths(-1).Month
    $monthName = (Get-Culture).DateTimeFormat.GetMonthName($month)
    $reportFolder = "$($script:CacheDataPath)\$($usageReport)\$($monthName)\$($inputReport)"
    Create-Directory $reportFolder

    # create a new DateTime object set to the first day of a given month and year
    $startOfMonth = Get-Date -Year $year -Month $month -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
    # add a month and subtract the smallest possible time unit
    $endOfMonth = ($startOfMonth).AddMonths(1).AddTicks(-1)    
    
    LogWrite -Message "Adding new column to csv files"
    #Invoke "AddColumntoCSVFile.ps1" - not implemented yet
    

    #----------------Read cert info-----------------------#            
    #$cert = Get-Item Cert:\LocalMachine\My\* | Where-Object { $_.Thumbprint -ieq "$($script:appThumbprintUsageReport)" }                
    #$script:authToken = Connect-NIHO365GraphWithCert -TenantName $script:TenantId -AppId $script:appIdOperationSupport -Certificate $cert    
    Invoke-GraphAPIAuthTokenCheck
    #----------------End-Read cert info-------------------#

    if ($script:authToken) {
        foreach ($report in $reportMetaData.UsageReport) {          
            $filename = "$($reportFolder)\$($Report["ReportName"]).csv"
            # check raw usage report file exists or not
            if (!(test-path $filename -PathType Leaf)) {		        
                if ($report["ReportEndpoint"] -ne ''){
                    #$ActivityResponse = (Get-NIHActivityReport -AuthToken $script:authToken -Report $report["ReportEndpoint"]) | ? { $_."Last Activity Date" -ne "" -and (Get-Date($_."Last Activity Date")) -igt $startOfMonth -and (Get-Date($_."Last Activity Date")) -ile $endOfMonth  } | Select * 
                    $ActivityResponse = (Get-NIHActivityReport -AuthToken $script:authToken -Report $report["ReportEndpoint"]) | ? { $_."Last Activity Date" -ne "" -and (Get-Date($_."Last Activity Date")) -igt $startOfMonth -and (Get-Date($_."Last Activity Date")) -ile $date  } | Select *
                    ExportCSV -DataSet $ActivityResponse -FileName $filename
                }
                #copy usage report file (has no endpoint) from Drop Folder to UsageReport folder
                #LogWrite -Message  "$($report["DropLocation"]) - $($reportFolder)"
                if ($report["ReportEndpoint"] -eq '' -and $report["DropLocation"] -ne ''){                
                    LogWrite -Message  "Copying files..$($report["DropLocation"]) - $($reportFolder)"
                    Copy-Item $report["DropLocation"] -Destination $reportFolder -Force
                }
	        }
            
        }
    }    
         
}

Try {
    #------------------Initialize global variables needed------------------#
    Set-LogFile -logFileName $logFileName
    Set-TenantVars
    Set-AzureAppVars
    Set-DataFile  
    #------------------End-Initialize global variables---------------------#
    
    LogWrite -Message "----------------------- [Populate Usage Reports Cache] Execution Started -----------------------"
    
    $script:StartTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    GetAllM365Reports
    $script:EndTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    LogWrite -Message "[Populate Usage Reports Cache] Start Time: $($script:StartTimeDailyCache)"
    LogWrite -Message "[Populate Usage Reports Cache] End Time:   $($script:EndTimeDailyCache)"

    LogWrite -Message  "----------------------- [Populate Usage Reports Cache] Execution Ended ------------------------"    
    
}
Catch [Exception] {
    LogWrite -Level ERROR "Error in the script: $($_)"    
}
