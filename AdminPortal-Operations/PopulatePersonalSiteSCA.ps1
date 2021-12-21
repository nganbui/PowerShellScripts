Param
(    
    [Parameter(Position = 0)]    
    [string]$CoIC
)

$0 = $MyInvocation.MyCommand.Definition
$dp0 = [System.IO.Path]::GetDirectoryName($0)
$currentFolder = $dp0.Substring($dp0.LastIndexOf('\') + 1)
$script:RootDir = Resolve-Path "$dp0\.." 
$logFileName = $MyInvocation.MyCommand.Name.Substring(0, $MyInvocation.MyCommand.Name.IndexOf('.'))
$logFileName = "$currentFolder\$logFileName"
$reportFile = "PersonalSiteSCA.csv"
#Include core
."$script:RootDir\Common\Lib\LibCore.ps1"
#Include dependent functionality
."$script:RootDir\Common\Lib\LibO365.ps1"
."$script:RootDir\Common\Lib\LibCache.ps1"
."$script:RootDir\Common\Lib\GraphAPILibSPOSites.ps1"
."$script:RootDir\Common\Lib\GraphAPILibSPOSitesDAO.ps1"
."$script:RootDir\Common\Lib\LibICs.ps1"
."$script:RootDir\Common\Lib\LibICsDAO.ps1"

<#
      ===========================================================================
      .DESCRIPTION
        Adding [ICName]_OD4BSecondaryAdmins and removing [OldICName]_OD4BSecondaryAdmins for Active Personal Sites
        Update SecondarySCA for PersonalSites table in DB
        - ICProfile DB Cache
        - PersonalSites DB Cache
        - Processing populate OD4B must get PersonalSites DB Cache because there is no IC info from O365 cache
#>

Function GeneratePersonSitesSCAReport {
    param($attachments)
    LogWrite -Message "Sending Email Report: [Populate Personal Site SCA]"    
    $subject = "[M365 DevOps] Populate IC-OneDrive Secondary Admin"
    $content = "There are currently not found any personal sites need to be populated SCA. Please review log file if there is any error."
    if ($attachments.Length -gt 0) { 
        $content = "<p><b>Description:</b> This job populate secondary admin for Personal Sites and update extended properties to DB.<br />"
        $content += "Please review and address any issues from the attached files if needed.</p>"    
    }
    $body = "<p><i>Note: This is an automated email. Please do not reply to this message.</i></p>
             $content
             <p>Script Start time: $($script:StartTimeOD4B)<br />
             Script End time: $($script:EndTimeOD4B)</p>
             <p>$($script:ProcessedICs)</p>
             <p>
             Total Active Personal Sites Retrieved and Processed: $($script:totalPersonalSitesRetrieved) <br />
             Total Active Personal Sites Records Udated: $($script:totalPersonalSitesUpdated) <br />
             Total Active Personal Sites Records UpdateFailed : $($script:totalPersonalSitesUpdateFailed)
             </p>
             <p><i>Thank you,</i> <br />NIH M365 Collaboration Support Team</p>"    
    #$body = [System.Web.HttpUtility]::HtmlDecode($body)

    SendEmail -subject $subject -body $body -Attachements $attachments #-To "ngan.bui@nih.gov"
    LogWrite -Message "Sending Email Report: [Populate SPO Site SCA] completed."
}

Function ResetPersonalConfig {
    param($FilePath)
    $icData = @(Import-csv -Path $FilePath)
    foreach ($item in $icData) {
        $item.Status = ""
        $item.CompletedDate = ""
    }
    $icData | Export-Csv $FilePath -NoTypeInformation

}

Try {
    #----------------Read PersonalSiteMetadata endpoint-----------------#            
    $path = "$dp0\PersonalSiteMetaData.psd1"
    if (Test-Path $path) {
        $psMetaData = Import-PowerShellDataFile -Path $path        
        if ($psMetaData) {
            $ICName = $psMetaData[[int]$CoIC]
            #$processedGroup = "Batch $CoIC"
        }        
    }
    else {
        $ICName = $null
    }

    <#
    LogWrite -Message "-Determine which week in the month..."
    $weekInMonth = (Get-WmiObject Win32_LocalTime).weekinmonth
    if ($weekInMonth  -eq 5)  { $weekInMonth  -=1 }        
    #>
    #$psMetaData[1]
    <#
    $d = Get-Date
    $e = [math]::Ceiling(($d.Day+(($d.AddDays(-($d.Day-1))).DayOfWeek.value__)-7)/7+1)
    if ($e -eq 5)  { $e -=1 }
        $e
    #>
            
    #----------------End-Read PersonalSiteMetadata endpoint-------------#
       
    #log file path
    Set-LogFile -logFileName $logFileName
    Set-DataFile
    Set-DBVars
    Set-TenantVars
    Set-AzureAppVars

    $script:StartTimeOD4B = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    LogWrite -Message "----------------------- [Populate Personal Site SCA] Execution Started -----------------------"
    #region Populate IC-OD4BSecondaryAdminsEmail for Personal Sites and update SecondarySCA for PersonalSites table in DB
    #LogWrite -Message "-Getting IC Data..."
    #SyncICProfileFromDBToCache
    $script:ICDBData = GetDataInCache -CacheType DB -ObjectType ICProfiles
    $script:ICDBData = $script:ICDBData | ? { $_.OD4BSecondaryAdminsEnabled -eq $true }
    $arrICs = @($script:ICDBData).ICName -split ","
         
    #----
    LogWrite -Message "-Getting PersonalSites Data..."
    SyncPersonalSitesFromDBToCache
    $script:PersonalSitesData = GetDataInCache -CacheType DB -ObjectType PersonalSites
    
    $psGroup = $script:PersonalSitesData | ? { $_.ICName -in $arrICs } | Group-Object -Property ICName | Sort-Object -Property Name
    #$psGroup = $script:PersonalSitesData | Group-Object -Property ICName | Sort-Object -Property Name    

    #----
    # Handling for single IC. If $ICName is passed, then the OD Secondary admins will be run alone for that IC. If $ICName equals Null, ICAdmins will be processed for all ICs    
    <#
    if ($psMetaData){
        $ICName = $psMetaData[[int]$weekInMonth]
        $processedGroup = "Batch $weekInMonth"
    }
    #>
    if ($ICName.Length) { 
        #$psGroup = @($psGroup | ? { $_.Name -eq $ICName })
        $arrICs = $ICName -split ","
        $script:ICDBData = $script:ICDBData.Where( { $_.ICName -in $arrICs })    
        $psGroup = @($psGroup | ? { $_.Name -in $arrICs })
    }

    #----        
    [System.Collections.ArrayList]$Report = @()

    $psGroup | & { process {
            $ICName = $_.Name
            $count = ($_.Group).Count
            LogWrite -Message "[$ICName] - Total: $count : Processing populate secondary admins personal sites..."
            $OD4BSecondaryAdminsEmail = ($script:ICDBData | ? { $_.ICName -eq $ICName }).OD4BSecondaryAdminsEmail.Trim()
            $OD4BSecondaryAdmins = ($script:ICDBData | ? { $_.ICName -eq $ICName }).OD4BSecondaryAdmins.Trim() 

            $_.Group | & { process {        
                    $psUrl = $_.URL
                    #$psUrl = "https://nih-my.sharepoint.com/personal/gordontr_nih_gov"          
                    $primarySCA = $_.PrimarySCA
                    #LogWrite -Message "Processing populate secondary admins [$OD4BSecondaryAdminsEmail] for personal sites [$psUrl]"
                
                    try {
                        LogWrite -Message "Connecting to personal site $psUrl ..."
                        $psConn = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $psURL                        
                        if ($null -ne $psConn) {
                            $personalSite = ValidateSite -Url $psURL -SiteContext $psConn
                            if ($personalSite.Status -eq 'Active' -and $personalSite.LockState -eq 'Unlock') {

                                $admins = Get-PnPSiteCollectionAdmin -Connection $psConn
                                if ($admins) {
                                    LogWrite -Message "Retrieving current personal site secondary admin from M365..."
                                    #$adminsEmail = ($admins).Email -join ";"
                                    $adminsUPN = ($admins).LoginName | % { $_.SubString($_.LastIndexOf("|") + 1) }
                    
                                    $arrSCAs = @($primarySCA, $OD4BSecondaryAdmins)
                                    $scaAdminsNotAllowed = ($admins.Where( { $_.LoginName.SubString($_.LoginName.LastIndexOf("|") + 1) -notin $arrSCAs })).LoginName
                                    #$scaAdminsRemoved = ($admins.Where({$_.LoginName.SubString($_.LoginName.LastIndexOf("|")+1) -notin $arrSCAs})).Email -join ";"
                    
                                    #--Adding correct <IC>_OD4BSecondaryAdminsEmails if not found from scaAdmins
                                    if ($adminsUPN -notcontains $OD4BSecondaryAdmins) {
                                        LogWrite -Message "-Adding [$OD4BSecondaryAdminsEmail]..."
                                        Add-PnPSiteCollectionAdmin -Owners $OD4BSecondaryAdminsEmail -Connection $psConn                           
                                    }
                                    #--Remove oldIC_OD4BSecondaryAdminsEmails if find wrong <IC>_OD4BSecondaryAdminsEmails
                                    if ($scaAdminsNotAllowed) {
                                        LogWrite -Message "-Removing [$scaAdminsNotAllowed]..."
                                        Remove-PnPSiteCollectionAdmin -Owners $scaAdminsNotAllowed -Connection $psConn
                                    }
                                }
                                LogWrite -Message "Validate if the personal site is updated after add/remove SCA"
                                #$personalSite = ValidateSite -Url $psURL -SiteContext $psConn
                                #if ($personalSite.Status -eq 'Active') {                      

                                LogWrite -Message "-Retrieving [SecondarySCA;FilesCount;Created;LastContentModifiedDate] ..."                                
                                try {                        
                                    $context = Get-PnPContext                
                                    $Web = $context.Web 
                                    $context.Load($Web) 
                                    $context.ExecuteQuery()
                                    
                                    $List = $context.Web.Lists.GetByTitle("Documents")
                                    $context.Load($List) 
                                    $context.ExecuteQuery()
                                    $FilesCount = $List.ItemCount
                                    $LastContentModifiedDate = $list.LastItemUserModifiedDate.toshortdatestring()                   
                                    #$NumberOfSubSites =  $Web.Webs.Count   # already handle in daily job          
                                    #$Description = $Web.Description  # already handle in daily job 
                                    $Created = $Web.Created.toshortdatestring()                   
                                    $SecondarySCA = (Get-PnPSiteCollectionAdmin | ? { $_.Email -ne '' -and $_.Email.ToLower() -notlike 'spoadm*' }).Email -join ";"                
                                    $context.Dispose()

                                    $siteObj = [PSCustomObject][ordered]@{
                                        ICName                  = $ICName
                                        SiteType                = "PersonalSites"
                                        URL                     = $psUrl
                                        SecondarySCA            = $SecondarySCA
                                        #WebsCount = $NumberOfSubSites
                                        FilesCount              = $FilesCount
                                        Created                 = $Created
                                        LastContentModifiedDate = $LastContentModifiedDate
                                        Operation               = "";
                                        OperationStatus         = ""; 
                                        AdditionalInfo          = ""
                                    }
                                    if ($siteObj) {
                                        #LogWrite -Message "-Updating [SecondarySCA;FilesCount;Created;LastContentModifiedDate] for [$psUrl] to DB..."
                                        #UpdateSPOSiteExtenedRecord -SqlConnection $script:connectionString -siteObj $siteObj
                                        $null = $Report.Add($siteObj)
                                    }

                                }
                                catch {
                                    LogWrite -Level ERROR -Message "An error occured processing personal site: $psUrl - $($_.Exception)" 

                                }
                            }
                        }
                    }
                    catch {
                        DisconnectPnpOnlineOAuth -Context $psConn
                        LogWrite -Level ERROR "-Unexpected Error [$psUrl]: $_ "                     
    
                    }
                    finally {
                        DisconnectPnpOnlineOAuth -Context $psConn
                    }
                }
            }
        } } 
       
    LogWrite -Message "[Completion] Update IC_OD4BSecondaryAdmins to personal sites in the tenant."
      
    $script:totalPersonalSitesRetrieved = 0
    $script:totalPersonalSitesUpdated = 0
    $script:totalPersonalSitesUpdateFailed = 0


    $attachedFiles = @()    
    if ($null -ne $Report -and $Report.Count -gt 0) {
        LogWrite -Message "-Updating [SecondarySCA;FilesCount;Created;LastContentModifiedDate] to DB..."        
        UpdateSitesExtenedInfoToDatabase $script:connectionString $Report
                            
        LogWrite -Message "Export to csv and sending report email to SP Admins..."
        $logPath = "$($script:DirLog)"
        if (!(Test-Path $logPath)) { 
            LogWrite -Message "Creating $logPath" 
            New-Item -ItemType "directory" -Path $logPath -Force
        }
        LogWrite -Message "Generating Log files..." 
        $reportFile = "$logPath\$reportFile"
                
        $Report | Export-Csv $reportFile -Encoding ASCII -NoTypeInformation
        $attachedFiles += $reportFile

        # Summary report
        #$psReport = $Report | Group-Object -Property ICName | Sort-Object -Property Name
        $psReport = $Report | Group-Object -Property ICName | Sort-Object -Property Name | Select Name, Count            
        $ProcessedICs = $psReport | ConvertTo-HTML -Fragment
        $ProcessedICs = [System.Web.HttpUtility]::HtmlDecode($ProcessedICs)
        $ProcessedICs = $ProcessedICs -replace '<table>', '<table cellpadding="5" cellspacing="2" style="border: 1px solid black;border-collapse: collapse;width:100%">'                    
        $ProcessedICs = $ProcessedICs -replace '<th>', '<th align="left" style="border: 1px solid black;">'
        $ProcessedICs = $ProcessedICs -replace '<td>', '<td style="border: 1px solid black;">'
        $script:ProcessedICs = $ProcessedICs

        $script:totalPersonalSitesRetrieved = $Report.Count
        $script:totalPersonalSitesUpdated = @($Report | Where-Object { ($_.OperationStatus -eq "Success") -and ($_.Operation -eq "Update") }).URL.Count
        $script:totalPersonalSitesUpdateFailed = @($Report | Where-Object { $_.OperationStatus -eq "Failed" }).URL.Count             
    }
    else {
        LogWrite -Message "There are currently no personal sites need to be populated SCA."        
    }

    $script:EndTimeOD4B = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    GeneratePersonSitesSCAReport -attachments $attachedFiles   
    
    LogWrite -Message "[Populate Personal Site SCA] Start Time: $($script:StartTimeOD4B)"
    LogWrite -Message "[Populate Personal Site SCA] End Time:   $($script:EndTimeOD4B)"
    LogWrite -Message  "----------------------- [Populate Personal Site SCA] Execution Ended ------------------------"    
    #endregion
    
}
Catch [Exception] {
    LogWrite -Level ERROR "-Unexpected Error: $_ "    
}
Finally {
    LogWrite -Message  "----------------------- [Populate Personal Site SCA] Completed ------------------------"
}