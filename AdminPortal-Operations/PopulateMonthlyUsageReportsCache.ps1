$0 = $MyInvocation.MyCommand.Definition
$dp0 = [System.IO.Path]::GetDirectoryName($0)
$currentFolder = $dp0.Substring($dp0.LastIndexOf('\')+1)
$script:RootDir = Resolve-Path "$dp0\.." 
$logFileName = $MyInvocation.MyCommand.Name.Substring(0, $MyInvocation.MyCommand.Name.IndexOf('.'))
$logFileName = "$currentFolder\$logFileName"
$usageReport = "UsageReports"
$inputReport = "Input"
$rawFolderName = "RawData"
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
    #$profileData = GetPsd1Data $path    
    #----------------End-Read report endpoint-------------#
    
    #--Create a folder UsageReports under Data if any        
    $date = Get-Date
    $year = $date.Year
    $month = $date.Month
    $month = $date.AddMonths(-1).Month
    $monthName = (Get-Culture).DateTimeFormat.GetMonthName($month)
    if (12 -eq $month){
        $year = $date.AddYears(-1).Year
    } 

    #$reportFolder = "$($script:CacheDataPath)\$($usageReport)\$($monthName)\$($inputReport)"
    LogWrite -Message  "Creating a directory name RawData if any..."
    $rawDataFolder = "$($script:CacheDataPath)\$usageReport\$year\$monthName\$rawFolderName"
    Create-Directory $rawDataFolder
    
    LogWrite -Message  "Creating a directory name Input if any..."
    $reportFolder = "$($script:CacheDataPath)\$usageReport\$year\$monthName\$inputReport"
    Create-Directory $reportFolder

    # create a new DateTime object set to the first day of a given month and year
    $startOfMonth = Get-Date -Year $year -Month $month -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
    # add a month and subtract the smallest possible time unit
    $endOfMonth = ($startOfMonth).AddMonths(1).AddTicks(-1)    
    
    #LogWrite -Message "Adding new column to csv files"
    #Invoke "AddColumntoCSVFile.ps1" - not implemented yet
    

    #----------------Read cert info-----------------------#                
    $cert = Get-Item Cert:\\LocalMachine\\My\* | Where-Object { $_.Subject -ieq "CN=$($script:appCertUsageReport)" }    
    $script:authToken = Connect-NIHO365GraphWithCert -TenantName $script:TenantName -AppId $script:appIdUsageReport -Certificate $cert    
    #Invoke-GraphAPIAuthTokenCheck
    #----------------End-Read cert info-------------------#

    if ($script:authToken) {
        foreach ($report in $reportMetaData.UsageReport) {          
            $rawDataFileName = "$($rawDataFolder)\$($Report["ReportName"]).csv"
            $filename = "$($reportFolder)\$($Report["ReportName"]).csv"
            # check raw usage report file exists or not
            if (!(test-path $filename -PathType Leaf)) {		        
                if ($report["ReportEndpoint"] -ne ''){
                    LogWrite -Message  "Downloading activity report to RawData folder..."                    
                    $ActivityRawData = (Get-NIHActivityReport -AuthToken $script:authToken -Report $report["ReportEndpoint"]) | Select *
                    if ($null -ne $ActivityRawData){
                        ExportCSV -DataSet $ActivityRawData -FileName $rawDataFileName
                    }
                    #---
                    #$ActivityResponse = (Get-NIHActivityReport -AuthToken $script:authToken -Report $report["ReportEndpoint"]) | ? { $_."Last Activity Date" -ne "" -and (Get-Date($_."Last Activity Date")) -igt $startOfMonth -and (Get-Date($_."Last Activity Date")) -ile $endOfMonth  } | Select * 
                    #$ActivityResponse = (Get-NIHActivityReport -AuthToken $script:authToken -Report $report["ReportEndpoint"]) | ? { $_."Last Activity Date" -ne "" -and (Get-Date($_."Last Activity Date")) -igt (Get-Date($startOfMonth)) -and (Get-Date($_."Last Activity Date")) -ile $date  } | Select *
                    
                    LogWrite -Message  "Filter out based on LastActivityDate before downloading activity report to Input folder..."
                    $ActivityResponse = (Get-NIHActivityReport -AuthToken $script:authToken -Report $report["ReportEndpoint"]) | ? { $_."Last Activity Date" -ne "" -and (Get-Date($_."Last Activity Date")) -ge (Get-Date($startOfMonth)) -and (Get-Date($_."Last Activity Date")) -le (Get-Date($endOfMonth))  } | Select *
                    #$ActivityResponse = (Get-NIHActivityReport -AuthToken $script:authToken -Report $report["ReportEndpoint"]) | Select *
                    if ($null -ne $ActivityResponse){
                        ExportCSV -DataSet $ActivityResponse -FileName $filename
                    }
                }
                #copy usage report file (has no endpoint) from Drop Folder to UsageReport folder                
                if ($report["ReportEndpoint"] -eq '' -and $report["DropLocation"] -ne ''){
                    if ([System.IO.File]::Exists($report["DropLocation"])){                        
                        Copy-Item $report["DropLocation"] -Destination $reportFolder -Force
                        LogWrite -Message  "Copy file $($report["DropLocation"]) to: $($reportFolder) completed."
                    }
                    else{
                        LogWrite -Message  "Report file: $($report["DropLocation"]) not found! Copying file skip."
                    }
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
