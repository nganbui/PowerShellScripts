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
."$script:RootDir\Common\Lib\GraphAPILibO365Users.ps1"
."$script:RootDir\Common\Lib\GraphAPILibO365UsersDAO.ps1"
."$script:RootDir\Common\Lib\LibCache.ps1"


Function GenerateM365UsersSyncReport {    
    #------
    LogWrite -Message "Sending Email Report: [Sync M365 Users]"    
    $subject = "[SPO-DevOps] M365 Users Daily Sync to DB"    
    $body = "<p><b>Description:</b> This job will sync the M365 Users to local database repository.  It updates all the users information regardless of changes to one or multiple fields. </p>"
    $body += "<p>Script Start time: $($script:StartTimeDailyCache)<br />"
    $body += "Script End time: $($script:EndTimeDailyCache)<br /><br />"    
    $body += GetM365UsersReportContent          
    SendEmail -subject $subject -body $body
    LogWrite -Message "Sending Email Report: [Sync M365 Users] completed."

}

Function GetM365UsersReportContent {
    $script:totalUsers = @($script:usersData).Count 
    $script:totalUsersAdded = @($script:usersData | Where-Object { ($_.OperationStatus -eq "Success") -and ($_.Operation -eq "Insert") }).Count
    $script:totalUsersUpdated = @($script:usersData | Where-Object { ($_.OperationStatus -eq "Success") -and ($_.Operation -eq "Update") }).Count    
    $script:totalUsersUpdateFailed = @($script:usersData | Where-Object { $_.OperationStatus -eq "Failed" }).Count
    $script:totalUsersWithSameSigninName = @($script:usersWithSameSigninName).Count

    $script:totalDelUsers = @($script:deletedUsersData).Count
    $script:totalDelUsersAdded = @($script:deletedUsersData | Where-Object { ($_.OperationStatus -eq "Success") -and ($_.Operation -eq "Insert") }).Count
    $script:totalDelUsersUpdated = @($script:deletedUsersData | Where-Object {($_.OperationStatus -eq "Success") -and ($_.Operation -eq "Update")}).Count
    $script:totalDelUsersUpdateFailed = @($script:deletedUsersData | Where-Object {$_.OperationStatus -eq "Failed"}).Count

    

    LogWrite -Message "Generating Email Report..."
    #----    
    LogWrite -Message "->Total Active Users Retrieved: $($script:totalUsers)"    
    LogWrite -Message "->Total User Records Inserted: $($script:totalUsersAdded)"
    LogWrite -Message "->Total User Records Updated: $($script:totalUsersUpdated)"
    LogWrite -Message "->Total User Records UpdateFailed: $($script:totalUsersUpdateFailed)" 
    LogWrite -Message "->Total number of users with same SigninName but different UserId: $($script:totalUsersWithSameSigninName)" 
    #---
    LogWrite -Message "->Total Deleted Users Retrieved: $($script:totalDelUsers)"
    LogWrite -Message "->Total Deleted Users Records Inserted: $($script:totalDelUsersAdded)"
    LogWrite -Message "->Total Deleted Users Updated: $($script:totalDelUsersUpdated)"
    LogWrite -Message "->Total Deleted Users UpdateFailed: $($script:totalDelUsersUpdateFailed)"     
        
    $msg = "<p>"
    $msg+="<p>Total Active Users Retrieved: $($script:totalUsers)<br />"
    $msg+="Total Deleted Users Retrieved: $($script:totalDelUsers)<br />" 
    $msg+= "=============================================================<br />"    
    $msg+="Total User Records Inserted: $($script:totalUsersAdded)<br />"
    $msg+="Total User Records Updated: $($script:totalUsersUpdated)<br />"
    $msg+="Total User Records Failed to Insert/Update: $($script:totalUsersUpdateFailed)<br/>"
    $msg+="Total number of users with same SigninName but different UserId: $($script:totalUsersWithSameSigninName)<br/>"
    $msg+= "=============================================================<br />" 
    $msg+="Total Deleted Users Records Inserted: $($script:totalDelUsersAdded)<br />"
    $msg+="Total Deleted Users Records Updated: $($script:totalDelUsersUpdated)<br />"
    $msg+="Total Deleted Users Records Failed to Insert/Update: $($script:totalDelUsersUpdateFailed)</p>"
    return $msg
}

Function GenerateM365UsersSyncLogs {    
    #$todaysDate = Get-Date -Format "MM-dd-yyyy"
    #$logPath = "$script:DirLog\$todaysDate"
    $logPath = "$($script:DirLog)"
    if (!(Test-Path $logPath)) { 
	    LogWrite -Message "Creating $logPath" 
        New-Item -ItemType "directory" -Path $logPath -Force
	} 

    $usersFile = "$logPath\O365ActiveUsers.csv"
    $delUsersFile = "$logPath\O365DeletedUsers.csv"
    $usersWithSameSigninFile = "$logPath\UsersWithSameSignin.csv"

    LogWrite -Message "Generating Log files..." 
    if ($script:usersData -and $script:usersData.Count -gt 0) {
        ExportCSV -DataSet $script:usersData -FileName $usersFile
    }
    if ($script:deletedUsersData  -and @($script:deletedUsersData).Count -gt 0) {
        ExportCSV -DataSet $script:deletedUsersData -FileName $delUsersFile
    }
    if ($script:usersWithSameSigninName -and $script:usersWithSameSigninName.Count -gt 0) {
        ExportCSV -DataSet $script:usersWithSameSigninName -FileName $usersWithSameSigninFile
    }
    
    LogWrite -Message "Generating Log files ended."     
}

Try {
    #log file path
    Set-LogFile -logFileName $logFileName    
    Set-DataFile
    $script:StartTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    LogWrite -Message "----------------------- [Sync M365 Users to DB] Execution Started --------------------------"
    #Get Users from Cache
    $script:usersData = @()
    $script:deletedUsersData = @()
    $script:usersWithSameSigninName = @()
    $script:guestsData = @()

    $script:usersData = GetDataInCache -CacheType O365 -ObjectType O365Users -ObjectState Active
    $script:deletedUsersData = GetDataInCache -CacheType O365 -ObjectType O365Users -ObjectState InActive  
    $script:guestsData =   GetDataInCache -CacheType O365 -ObjectType O365Guests -ObjectState Active  

    #Members
    if ($script:usersData -eq $null) {
        LogWrite -Message "M365 Users data not found in cache. Processing from M365"
        #Retrieve O365 Users...
        Set-TenantVars
        Set-AzureAppVars        
        GetAllM365Users
        #Cache Users
        CacheO365Users
        $script:usersData = GetDataInCache -CacheType O365 -ObjectType O365Users -ObjectState Active
        $script:deletedUsersData = GetDataInCache -CacheType O365 -ObjectType O365Users -ObjectState InActive         
    }
    else {
        LogWrite -Message "Processing M365 Users data from cache"
    }
    #Guests
    if ($script:guestsData -eq $null) {
        LogWrite -Message "Guests data not found in cache. Processing from M365"
        #Retrieve O365 Users...
        Set-TenantVars
        Set-AzureAppVars        
        GetGuestUsers
        #Cache Users
        CacheO365Users
        $script:guestsData =   GetDataInCache -CacheType O365 -ObjectType O365Guests -ObjectState Active
    }
    else {
        LogWrite -Message "Processing M365 Users data from cache"
    }
    Set-DBVars
    #Update All Users to Database
    UpdateO365UsersToDatabase
    $script:EndTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"  
    #Generate Log files and send Email Report
    Set-EmailVars
    GenerateM365UsersSyncReport
    #Generate Log files
    GenerateM365UsersSyncLogs
   
    
    LogWrite -Message "[Sync M365 Users to DB] Start Time: $($script:StartTimeDailyCache)"
    LogWrite -Message "[Sync M365 Users to DB] End Time:   $($script:EndTimeDailyCache)"
    LogWrite -Message  "----------------------- [Sync M365 Users to DB] Execution Ended --------------------------"    
    
}
Catch [Exception] {
    LogWrite -Level ERROR "-Unexpected Error: $_ "
}
Finally {
    LogWrite -Message  "----------------------- [Sync M365 Users to DB] Completed ------------------------"
}