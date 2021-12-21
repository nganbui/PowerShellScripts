$reportFolder = "D:\Reports"
$files = @('TeamsUsageActivity.csv','PSTNCallUsage.csv')

$date = Get-Date
$year = $date.Year
$month = $date.AddMonths(-1).Month
$startOfMonth = Get-Date -Year $year -Month $month -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
# add a month and subtract the smallest possible time unit
$endOfMonth = ($startOfMonth).AddMonths(1).AddTicks(-1) 

  
foreach($f in $files){
    $outfile = "$reportFolder\$f"    
    $csv = Import-Csv $outfile
    $newcsv = @()
    $refreshDate
    $reportPeriod = $null

    $headers = ($csv | Get-Member -MemberType NoteProperty).Name

    if('Report Refresh Date' -notin $headers){
        $refreshDate =  $endOfMonth
    }
    if('Report Period' -notin $headers){
        $reportPeriod =  "30"
    }
    if ($refreshDate -eq $null -and $reportPeriod -eq $null){
        continue
    }
    foreach ($row in $csv ) {
        if($refreshDate -ne $null){
            $row | Add-Member -MemberType NoteProperty -Name 'Report Refresh Date' -Value $endOfMonth
        }
        if($reportPeriod -ne $null){
            $row | Add-Member -MemberType NoteProperty -Name 'Report Period' -Value $reportPeriod
        }
        $newcsv += $row
    }
    
    if ($f -eq "TeamsUsageActivity.csv"){        
        $newcsv | Select 'Report Refresh Date',Id,DisplayName,Privacy,ActiveUsers,ActiveChannels,Guests,ReplyMessages,PostMessages,MeetingsOrganized,UrgentMessages,Reactions,Mentions,ChannelMessages,        
        @{ Name = 'LastActivityDate'; Expression = { $_.'LastActivity (UTC Time)' } },'Report Period'  | Export-CSV $outfile -Encoding ascii -NoTypeInformation -Force       
    }
    if ($f -eq "PSTNCallUsage.csv"){
        $newcsv | Select 'Report Refresh Date',UsageId,'Call ID','Conference ID','User Location','AAD ObjectId',UPN,'User Display Name','Caller ID','Call Type','Number Type',
        @{ Name = 'DomesticInternational'; Expression = { $_.'Domestic/International' } },'Destination Dialed','Destination Number','Start Time','End Time','Duration Seconds','Connection Fee',Charge,Currency,Capability,'Report Period' | Export-CSV $outfile -Encoding ascii -NoTypeInformation -Force
    }

    #$newcsv | Select 'Report Refresh Date',UPN,Email,Name,'Last Password Change','Creation Type','User State',StateChangeDateTime,'Age (Days)',Created,'Last Login Date',Operation,SiteUrl,SourceFileName,GroupIds,GroupNames,ClientIP,UserAgent,DN,'Report Period' | Export-CSV $outfile -Encoding ascii -NoTypeInformation
    #$newcsv | Select 'Report Refresh Date',Id, DisplayName, Privacy, ActiveUsers, ActiveChannels, Guests, ReplyMessages, PostMessages, MeetingsOrganized, UrgentMessages, Reactions, Mentions, ChannelMessages, LastActivityDate,'Report Period' | Export-CSV $outfile -Encoding ascii -NoTypeInformation
    #$newcsv | Select 'Report Refresh Date',UsageId,CallID,ConferenceID,UserLocation,AADObjectId,UPN,UserDisplayName,CallerID,CallType,NumberType,Domestic/International,DestinationDialed,DestinationNumber,StartTime,EndTime,DurationSeconds,ConnectionFee,Charge,Currency,Capability,'Report Period' | Export-CSV $outfile -Encoding ascii -NoTypeInformation
    
    #$newcsv | Select 'Report Refresh Date',UsageId,'Call ID','Conference ID','User Location','AAD ObjectId',UPN,'User Display Name','Caller ID','Call Type','Number Type',
    #@{ Name = 'DomesticInternational'; Expression = { $_.'Domestic/International' } },'Destination Dialed','Destination Number','Start Time','End Time','Duration Seconds','Connection Fee',Charge,Currency,Capability,'Report Period' | Export-CSV $outfile -Encoding ascii -NoTypeInformation
}