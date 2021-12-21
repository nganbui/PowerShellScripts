$0 = $MyInvocation.MyCommand.Definition
$dp0 = [System.IO.Path]::GetDirectoryName($0)
$currentFolder = $dp0.Substring($dp0.LastIndexOf('\')+1)
$script:RootDir = Resolve-Path "$dp0\.." 
$logFileName = $MyInvocation.MyCommand.Name.Substring(0, $MyInvocation.MyCommand.Name.IndexOf('.'))
$logFileName = "$currentFolder\$logFileName"
$usageReport = "UsageReports"
$logSyncFileName = "SyncDB.txt"
$inputReport = "Input"
#Include core
."$script:RootDir\Common\Lib\LibCore.ps1"
."$script:RootDir\Common\Lib\LibReportsDAO.ps1"
#Include dependent functionality

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


function Out-DataTable { 
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
                    $DR.Item($property.Name) = $null                    
                    if($property.value -ne $null -and $property.value -ne ''){
                        $DR.Item($property.Name) = $property.value
                    } 
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


Try {
    #---log file path
    Set-LogFile -logFileName $logFileName
    $script:StartTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    LogWrite -Message "----------------------- [Sync Usage Reports Execution Started --------------------------"    
    #---DB and Data configuration
    Set-DBVars
    Set-DataFile 
    #---Build the sqlbulkcopy connection, and set the timeout to infinite
    $cn = new-object System.Data.SqlClient.SqlConnection("$($script:ConnectionString)")
    $cn.Open()
    $bc = new-object ("System.Data.SqlClient.SqlBulkCopy") $cn
    $bc.BatchSize = 50000
    $bc.BulkCopyTimeout = 0
    #---- Usage Reports for previous month    
    $date = Get-Date
    $year = $date.Year
    $month = $date.Month
    $month = $date.AddMonths(-1).Month
    $monthName = (Get-Culture).DateTimeFormat.GetMonthName($month)
    if (12 -eq $month){
        $year = $date.AddYears(-1).Year
    } 
    $reportFolder = "$($script:CacheDataPath)\$usageReport\$year\$monthName\$inputReport"    
    $reportfiles = Get-ChildItem $reportFolder -Filter "*.csv"
    
    # create a .txt file to capture usage report have synced to DB if any
    $synFilename = "$($reportFolder)\$logSyncFileName"
    if (![System.IO.File]::Exists($synFilename)) {
        New-Item -ItemType File -Path $synFilename -Force        
    }    
    LogWrite -Message "Wipe TeamsOnlyUsers before sync to DB..."
    DelTeamsOnlyUsers $script:connectionString
    
    foreach($file in $reportfiles ){       
       $foundBaseName = Get-Content $synFilename | Where-Object { $_.Contains($file.BaseName) }
       #$foundBaseName = @(Get-Content $synFilename | Where-Object { $_.Contains($file.BaseName) }).Count
       # not synced to DB yet then sync + update sync log file
       if (!$foundBaseName){
            $data = Import-Csv $file.FullName | Out-DataTable
            $data.CaseSensitive = $false
            $bc.DestinationTableName = $file.BaseName            
            <#foreach ($col in $data.Columns) {
                $bc.ColumnMappings.Add($col.ColumnName, $col.ColumnName)
                LogWrite -Message "$col"
            }#>
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
Catch [Exception] {
    LogWrite -Level ERROR "Error in the script: $($_)" 
}
