[cmdletBinding()]
    param(
        [parameter(Mandatory = $false)]
        [String]$FromDate=$null,        
        [parameter(Mandatory = $false)]        
        [String]$ToDate=$null       
    )   

$0 = $MyInvocation.MyCommand.Definition
$dp0 = [System.IO.Path]::GetDirectoryName($0)
$currentFolder = $dp0.Substring($dp0.LastIndexOf('\')+1)
$script:RootDir = Resolve-Path "$dp0\.." 
$logFileName = $MyInvocation.MyCommand.Name.Substring(0, $MyInvocation.MyCommand.Name.IndexOf('.'))
$logFileName = "$currentFolder\$logFileName"
$usageReport = "UsageReports"
$logSyncFileName = "SyncDB.txt"
$inputReport = "Input"
$outpuReport = "Output"
#Include core
."$script:RootDir\Common\Lib\LibCore.ps1"
#Include dependent functionality
."$script:RootDir\Common\Lib\LibReportsDAO.ps1"
."$script:RootDir\Common\Lib\LibO365.ps1"
."$script:RootDir\Common\Lib\LibCache.ps1"

<# 
    The report should run on the second day of a month
    Reports are available for the last 7 days, 30 days, 90 days, and 180 days. 
    Data won't exist for all reporting periods right away. 
    The reports become available within 48 hours.
#>

Function Get-Type { 
    param($type) 
 
$types = @( 
'System.Boolean', 
'System.Byte[]', 
'System.Byte', 
'System.Char', 
'System.Datetime', 
'System.Decimal', 
'System.Double', 
'System.Guid', 
'System.Int16', 
'System.Int32', 
'System.Int64', 
'System.Single', 
'System.UInt16', 
'System.UInt32', 
'System.UInt64') 
 
    if ( $types -contains $type ) { 
        Write-Output "$type" 
    } 
    else { 
        Write-Output 'System.String' 
         
    } 
}

Function Out-DataTable { 
    [CmdletBinding()] 
    param([Parameter(Position=0, Mandatory=$true, ValueFromPipeline = $true)] [PSObject[]]$InputObject) 
 
    Begin 
    { 
        $dt = new-object Data.datatable   
        $First = $true  
    } 
    Process 
    { 
        foreach ($object in $InputObject) 
        { 
            $DR = $DT.NewRow()   
            foreach($property in $object.PsObject.get_properties()) 
            {   
                if ($first) 
                {   
                    $Col =  new-object Data.DataColumn   
                    $Col.ColumnName = $property.Name.ToString()   
                    if ($property.value) 
                    { 
                        if ($property.value -isnot [System.DBNull]) { 
                            $Col.DataType = [System.Type]::GetType("$(Get-Type $property.TypeNameOfValue)") 
                         } 
                    } 
                    $DT.Columns.Add($Col) 
                }   
                if ($property.Gettype().IsArray) { 
                    $DR.Item($property.Name) =$property.value | ConvertTo-XML -AS String -NoTypeInformation -Depth 1 
                }   
               else { 
                    $DR.Item($property.Name) = $property.value 
                } 
            }   
            $DT.Rows.Add($DR)   
            $First = $false 
        } 
    }  
      
    End 
    { 
        Write-Output @(,($dt)) 
    } 
 
}

Function SetUsageReportsVar{
    # Create a folder UsageReports under Data if any        
    $date = Get-Date
    $year = $date.Year
    $month = $date.AddMonths(-1).Month
    $monthName = (Get-Culture).DateTimeFormat.GetMonthName($month)
    $script:ReportInput = "$($script:CacheDataPath)\$($usageReport)\$($monthName)\$($inputReport)"
    Create-Directory $script:ReportInput
    $script:ReportOutput = "$($script:CacheDataPath)\$($usageReport)\$($monthName)\$($outpuReport)"    
    Create-Directory $script:ReportOutput

    # create a new DateTime object set to the first day of a given month and year
    $startOfMonth = Get-Date -Year $year -Month $month -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
    # add a month and subtract the smallest possible time unit
    $endOfMonth = ($startOfMonth).AddMonths(1).AddTicks(-1)
    # By default value previous month
    $script:StartDate = $startOfMonth.ToString("MM/dd/yyyy")
    #$script:EndDate = $endOfMonth.ToString("MM/dd/yyyy")
    $script:EndDate = $date.ToString("MM/dd/yyyy")

    # Use $FromDate & $ToDate if they are not empty    
    if(-not [string]::IsNullOrEmpty($FromDate) -and (-not [string]::IsNullOrEmpty($ToDate)) ){
        $script:StartDate = $FromDate
        $script:EndDate = $ToDate 
        $script:ReportInput = "$($script:CacheDataPath)\$($usageReport)\$($monthName)\$($date)\$($inputReport)"
        Create-Directory $script:ReportInput
        $script:ReportOutput = "$($script:CacheDataPath)\$($usageReport)\$($monthName)\$($date)\$($outpuReport)"    
        Create-Directory $script:ReportOutput
 
    }    
    LogWrite -Message "----------------------- Generate Usage Reports [$($script:StartDate)] - [$($script:EndDate)] -----------------------"
}

Function PullUsageReports {
    #----------------Read report endpoint-----------------#            
    $path = "$dp0\ReportsMetadata.psd1"
    $reportMetaData = Import-PowerShellDataFile -Path $path
    #----------------End-Read report endpoint-------------#
    #----------------Read cert info-----------------------#            
    #$cert = Get-Item Cert:\LocalMachine\My\* | Where-Object { $_.Thumbprint -ieq "$($script:appThumbprintUsageReport)" }                
    #$script:authToken = Connect-NIHO365GraphWithCert -TenantName $script:TenantId -AppId $script:appIdOperationSupport -Certificate $cert    
    Invoke-GraphAPIAuthTokenCheck #APP ID AND SECRET
    #----------------End-Read cert info-------------------#    
    if ($script:authToken) {
        foreach ($report in $reportMetaData.UsageReport) {          
            $filename = "$($script:ReportInput)\$($Report["ReportName"]).csv"
            # check raw usage report file exists or not
            if (!(test-path $filename -PathType Leaf)) {		        
                if ($report["ReportEndpoint"] -ne ''){
                    #$ActivityResponse = (Get-NIHActivityReport -AuthToken $script:authToken -Report $report["ReportEndpoint"]) | ? { $_."Last Activity Date" -ne "" -and (Get-Date($_."Last Activity Date")) -igt $startOfMonth -and (Get-Date($_."Last Activity Date")) -ile $endOfMonth  } | Select *           
                    #$ActivityResponse = (Get-NIHActivityReport -AuthToken $script:authToken -Report $report["ReportEndpoint"]) | ? { $_."Last Activity Date" -ne "" -and (Get-Date($_."Last Activity Date")) -igt $($script:StartDate) -and (Get-Date($_."Last Activity Date")) -ile $($script:EndDate)  } | Select *
                    $ActivityResponse = (Get-NIHActivityReport -AuthToken $script:authToken -Report $report["ReportEndpoint"]) | ? { $_."Last Activity Date" -ne "" -and (Get-Date($_."Last Activity Date")) -igt $($script:StartDate) -and (Get-Date($_."Last Activity Date")) -ile $($script:EndDate)  } | Select *
                    ExportCSV -DataSet $ActivityResponse -FileName $filename
                }
                #copy usage report file (has no endpoint) from Drop Folder to UsageReport folder
                elseif ($report["ReportEndpoint"] -eq '' -and $report["DropLocation"] -ne ''){                
                    Copy-Item $report["DropLocation"] -Destination $script:ReportInput -Force
                }
	        }
            
        }
    }    
         
}

Function SyncUsageReports {    
    $script:StartTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    LogWrite -Message "----------------------- [Sync Usage Reports Execution Started --------------------------"    
    
    #---Build the sqlbulkcopy connection, and set the timeout to infinite
    $cn = new-object System.Data.SqlClient.SqlConnection("$($script:ConnectionString)")
    $cn.Open()
    $bc = new-object ("System.Data.SqlClient.SqlBulkCopy") $cn
    $bc.BatchSize = 50000
    $bc.BulkCopyTimeout = 0
    #---- Usage Reports for previous month    
    $reportfiles = Get-ChildItem $script:ReportInput -Filter "*.csv"
    # create a .txt file to capture usage report have synced to DB if any
    $synFilename = "$($script:ReportInput)\$logSyncFileName"
    if (![System.IO.File]::Exists($synFilename)) {
        New-Item -ItemType File -Path $synFilename -Force        
    }
    foreach($file in $reportfiles ){       
       $foundBaseName = Get-Content $synFilename | Where-Object { $_.Contains($file.BaseName) }
       #$foundBaseName = @(Get-Content $synFilename | Where-Object { $_.Contains($file.BaseName) }).Count

       # not synced to DB yet then sync + update sync log file
       if (!$foundBaseName){
            $data = Import-Csv $file.FullName | Out-DataTable
            $bc.DestinationTableName = $file.BaseName
            $bc.WriteToServer($data)
            $data.Dispose()
            Add-Content -Path $synFilename -Value "`n$($file.BaseName)" -Force
       }
    }
    $cn.Close()
    $cn.Dispose()
    $bc.Dispose()
   
    $script:EndTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"   
    LogWrite -Message "[Sync Usage Reports] Start Time: $($script:StartTimeDailyCache)"
    LogWrite -Message "[Sync Usage Reports] End Time:   $($script:EndTimeDailyCache)"
    LogWrite -Message  "----------------------- [Sync Usage Reports] Execution Ended --------------------------"      
         
}

Function GenerateUsageReports {
    #----------------Read report endpoint-----------------#            
    $path = "$dp0\ReportsMetadata.psd1"
    $reportMetaData = Import-PowerShellDataFile -Path $path
    #----------------End-Read report endpoint-------------#

    foreach ($report in $reportMetaData.ReportConfig) { 
        ## Create Arrays  
        $results = @()    
        $queryResults = GetReports -connectionString $script:connectionString -StoredProcedureName $report["StoredProc"] -StartDate $($script:StartDate) -EndDate $($script:EndDate)
        foreach($o in $queryResults) { 
            if ($o.ICName){
                $results += $o  
            }
        }
        $filename = "$($script:ReportOutput)\$($report["FileName"]).csv"       
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
    Set-TenantVars
    Set-AzureAppVars
    Set-DataFile      
    Set-DBVars     
    #------------------End-Initialize global variables---------------------#
    
    LogWrite -Message "----------------------- [Generate Usage Reports] Execution Started -----------------------"
    #------------------ UsageReport variables---------------------#
    SetUsageReportsVar
    #------------------ End-UsageReport variables-----------------#    

    $script:StartTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss" 
    PullUsageReports
    SyncUsageReports          
    GenerateUsageReports
    $script:EndTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    LogWrite -Message "[Generate Usage Reports] Start Time: $($script:StartTimeDailyCache)"
    LogWrite -Message "[Generate Usage Reports] End Time:   $($script:EndTimeDailyCache)"

    LogWrite -Message  "----------------------- [Generate Usage Reports] Execution Ended ------------------------"    
    
}
Catch [Exception] {
    LogWrite -Level ERROR "Error in the script: $($_)"    
}
