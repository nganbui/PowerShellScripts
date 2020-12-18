$0 = $MyInvocation.MyCommand.Definition
$script:dp0 = [System.IO.Path]::GetDirectoryName($0)
$script:RootDir = Resolve-Path "$script:dp0\.." 
$logFileName = $MyInvocation.MyCommand.Name.Substring(0, $MyInvocation.MyCommand.Name.IndexOf('.'))
$usageReport = "Other"
#Include core
."$script:RootDir\Common\Lib\LibCore.ps1"
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


Try {
    #---log file path
    Set-LogFile -logFileName $logFileName
    $script:StartTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    LogWrite -Message "----------------------- [Populate TeamsOnly Users Execution Started --------------------------"    
    #---DB and Data configuration
    Set-DBVars
    Set-DataFile 
    #---Build the sqlbulkcopy connection, and set the timeout to infinite
    $cn = new-object System.Data.SqlClient.SqlConnection("$($script:ConnectionString)")
    $cn.Open()
    $bc = new-object ("System.Data.SqlClient.SqlBulkCopy") $cn
    $bc.BatchSize = 50000
    $bc.BulkCopyTimeout = 0
    $reportFolder = "$($script:CacheDataPath)\$($usageReport)"
    $reportfiles = Get-ChildItem $reportFolder -Filter "*.csv"
    
    foreach($file in $reportfiles ){       
        $data = Import-Csv $file.FullName | Out-DataTable
        $bc.DestinationTableName = $file.BaseName
        $bc.WriteToServer($data)
        $data.Dispose()
    }
    $cn.Close()
    $cn.Dispose()
    $bc.Dispose()
   
    $script:EndTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"   
    LogWrite -Message "[Populate TeamsOnly Users] Start Time: $($script:StartTimeDailyCache)"
    LogWrite -Message "[Populate TeamsOnly Users] End Time:   $($script:EndTimeDailyCache)"
    LogWrite -Message  "----------------------- [Sync Usage Reports] Execution Ended --------------------------"    
    
}
Catch [Exception] {
    LogWrite -Level ERROR "Error in the script: $($_)" 
}
