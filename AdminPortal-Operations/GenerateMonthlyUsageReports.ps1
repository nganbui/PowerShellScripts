$0 = $MyInvocation.MyCommand.Definition
$dp0 = [System.IO.Path]::GetDirectoryName($0)
$currentFolder = $dp0.Substring($dp0.LastIndexOf('\')+1)
$script:RootDir = Resolve-Path "$dp0\.." 
$logFileName = $MyInvocation.MyCommand.Name.Substring(0, $MyInvocation.MyCommand.Name.IndexOf('.'))
$logFileName = "$currentFolder\$logFileName"
$usageReport = "UsageReports"
$outpuReport = "Output"
#Include core
."$script:RootDir\Common\Lib\LibCore.ps1"
#Include dependent functionality
."$script:RootDir\Common\Lib\LibReportsDAO.ps1"


Function GenerateUsageReports {    
    #--Create a folder UsageReports under Data if any        
    $date = Get-Date
    $year = $date.Year
    $month = $date.AddMonths(-1).Month
    $monthName = (Get-Culture).DateTimeFormat.GetMonthName($month)
    $reportFolder = "$($script:CacheDataPath)\$($usageReport)\$($monthName)\$($outpuReport)"
    Create-Directory $reportFolder
    # create a new DateTime object set to the first day of a given month and year
    $startOfMonth = Get-Date -Year $year -Month $month -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
    # add a month and subtract the smallest possible time unit
    $endOfMonth = ($startOfMonth).AddMonths(1).AddTicks(-1)

    $startOfMonth = $startOfMonth.ToString("MM/dd/yyyy")
    $endOfMonth = $endOfMonth.ToString("MM/dd/yyyy")

    Set-DBVars
    #----------------Read report endpoint-----------------#            
    $path = "$dp0\ReportsMetadata.psd1"
    $reportMetaData = Import-PowerShellDataFile -Path $path
    #----------------End-Read report endpoint-------------#
    $i = 1    
    foreach ($report in $reportMetaData.ReportConfig) {
        $storedName =  $report["StoredProc"]        
        $filename = "$($reportFolder)\$($report["FileName"]).csv"
        $baseline = $report["Baseline"]
        ## Create Arrays  
        $results = @()
        <#if ($report["Baseline"] -eq $null){  
            $queryResults = GetReports -connectionString $script:connectionString -StoredProcedureName $report["StoredProc"] -StartDate $startOfMonth -EndDate $endOfMonth
        }
        elseif ($report["Baseline"] -eq "Yes"){  
            $queryResults = GetReports -connectionString $script:connectionString -StoredProcedureName $report["StoredProc"] -Baseline
        }#>
        
        LogWrite -Message "$($i).: Generating report: $($report["FileName"])" 
        if ($baseline -eq "Yes"){  
            $queryResults = GetBaselineReports -connectionString $script:connectionString -StoredProcedureName $storedName
        }
        else{
            $queryResults = GetReports -connectionString $script:connectionString -StoredProcedureName $storedName -StartDate $startOfMonth -EndDate $endOfMonth
        }
        
        if ($queryResults.Count -gt 0){
            foreach($o in $queryResults) { 
                if ($o.ICName){
                    $results += $o  
                }
                else{
                    if ($o.UserPrincipalName){
                        $results += $o  
                    }
                }
            }
        }
        $i++               
        ExportCSV -DataSet $results -FileName $filename 
               
    }
}

Function Import-Excel {
    param (
    [string]$FileName,
    [string]$WorksheetName
    )

    if ($FileName -eq "") {
        throw "Please provide path to the Excel file"
        Exit
    }

    if (-not (Test-Path $FileName)) {
        throw "Path '$FileName' does not exist."
        Exit
    }

    $strSheetName = $WorksheetName + '$'
    $query = 'select * from ['+$strSheetName+']';

    $connectionString = 
      "Provider=Microsoft.ACE.OLEDB.12.0;"
    
}

Try {
    #------------------Initialize global variables needed------------------#
    Set-LogFile -logFileName $logFileName    
    Set-DataFile  
    #------------------End-Initialize global variables---------------------#
    
    LogWrite -Message "----------------------- [Generate Usage Reports] Execution Started -----------------------"
    
    $script:StartTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    GenerateUsageReports
    $script:EndTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    LogWrite -Message "[Generate Usage Reports] Start Time: $($script:StartTimeDailyCache)"
    LogWrite -Message "[Generate Usage Reports] End Time:   $($script:EndTimeDailyCache)"

    LogWrite -Message  "----------------------- [Generate Usage Reports] Execution Ended ------------------------"    
    
}
Catch [Exception] {
    LogWrite -Level ERROR "Error in the script: $($_)"    
}
