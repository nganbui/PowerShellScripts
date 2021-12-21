Param
(    
    [Parameter(Mandatory = $false)]
    [string]$ICName = "FIC"
)

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
."$script:RootDir\Common\Lib\LibCache.ps1"
."$script:RootDir\Common\Lib\GraphAPILibSPOSites.ps1"
."$script:RootDir\Common\Lib\GraphAPILibSPOSitesDAO.ps1"
."$script:RootDir\Common\Lib\LibICs.ps1"
."$script:RootDir\Common\Lib\LibICsDAO.ps1"


<#
      ===========================================================================
      .DESCRIPTION
        Adding [ICName]_SPOSecAdmins as SCA for SPO Sites
        Update SecondarySCA for SPO Sites in DB
        - ICProfile DB Cache
        - Sites DB Cache
        - Processing populate [ICName]_SPOSecAdmins must get Sites DB Cache because there is no IC info from O365 cache                     
    #>

Function GenerateSiteSCAReport{
    param($attachments)
    LogWrite -Message "Sending Email Report: [Populate SPO Site SCA]"    
    $subject = "[M365 DevOps] Populate IC-SPO Secondary Admin"
    $content = "There are currently not found any sites need to be populated SCA. Please review log file if there is any error."
    if ($attachments.Length -gt 0) { 
        $content = "<p><b>Description:</b> This job populate secondary admin for sharepoint sites.<br />"
        $content += "Please review and address any issues from the attached files if needed.</p>"    
    }
    $body = "<p><i>Note: This is an automated email. Please do not reply to this message.</i></p>
             $content
             <p><i>Thank you,</i> <br />NIH M365 Collaboration Support Team</p>"    
    $body = [System.Web.HttpUtility]::HtmlDecode($body)

    SendEmail -subject $subject -body $body -Attachements $attachments #-To "ngan.bui@nih.gov"
    LogWrite -Message "Sending Email Report: [Populate SPO Site SCA] completed."
}

Try {
    #log file path
    Set-LogFile -logFileName $logFileName
    Set-DataFile
    Set-DBVars
    Set-TenantVars
    Set-AzureAppVars
    $script:StartTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"           
    LogWrite -Message "----------------------- [Populate SPO Site SCA] Execution Started -----------------------"
    
    $spoAdmSvc = "SPOADMSVC@nih.gov"
    $cloudSvc = "m365collabadmsvc@nih.onmicrosoft.com"

    #region Populate IC-OD4BSecondaryAdminsEmail for Personal Sites and update SecondarySCA for PersonalSites table in DB
    LogWrite -Message "-Getting IC Data..."
    SyncICProfileFromDBToCache
    $script:ICDBData = GetDataInCache -CacheType DB -ObjectType ICProfiles
    #----
    LogWrite -Message "-Getting Sites Data..."
    SyncSitesFromDBToCache
    $script:SitesData = GetDataInCache -CacheType DB -ObjectType SPOSites
    
    #-populate IC secondary admin to all sites except Private Channel site
    #$psGroup = $script:SitesData | ? {$_.LockState -eq 'Unlock' -and $_.TemplateId -ne 'TEAMCHANNEL#0' -and $_.ICName -ne '' -and $_.ICName -ne 'Other'} | Group-Object -Property ICName | Sort-Object -Property Name    
    $psGroup = $script:SitesData | ? {$_.LockState -eq 'Unlock' -and $_.ICName -ne '' -and $_.ICName -ne 'Other'} | Group-Object -Property ICName | Sort-Object -Property Name    
    #----
    # Handling for single IC. If $ICName is passed, then the Secondary admins will be run alone for that IC. If $ICName equals Null, ICAdmins will be processed for all ICs
    if ($null -ne $ICName) { 
        $psGroup = @($psGroup | ? { $_.Name -eq $ICName })
    }    
    #----        
    [System.Collections.ArrayList]$Report = @()

    $psGroup | & { process {
        $ICName = $_.Name
        $count = ($_.Group).Count
        LogWrite -Message "Processing populate secondary admins for sites for IC <$ICName> - Total sites: $count"
        $SPOSecondaryAdminsEmail = ($script:ICDBData | ? { $_.ICName -eq $ICName}).SPOSecondaryAdminsEmail.Trim()
        $SPOSecondaryAdmins = ($script:ICDBData | ? { $_.ICName -eq $ICName}).SPOSecondaryAdmins.Trim() 

        #LogWrite -Message "-Filter out other IC Second Admins..."
        $scaAdminsNotAllowed = @($script:ICDBData | ? { $_.ICName -ne $ICName -and $_.SPOSecondaryAdmins -ne ''}).SPOSecondaryAdmins.Trim()        
        #$scaAdminsNotAllowed = $scaAdminsNotAllowed.Substring(0,$scaAdminsNotAllowed.Length-1)

        $_.Group | & { process {
            $psUrl = $_.URL            
            $primarySCA =$_.PrimarySCA
            LogWrite -Message "Processing populate secondary admins [$SPOSecondaryAdminsEmail] for site [$psUrl]"
            try{
                $psConn = ConnectPnpOnlineOAuth -TenantId $script:TenantId -ClientId $script:appIdOperationSupport -Thumbprint $script:appThumbprintOperationSupport -Url $psURL                        
                $admins = Get-PnPSiteCollectionAdmin -Connection $psConn
                $adminsEmail = ($admins).Email -join ";"
                $adminsUPN = ($admins).LoginName | % {$_.SubString($_.LastIndexOf("|")+1)}
                $updateToDB = $false
                
                LogWrite -Message "Retrieving current site secondary admin from M365..."

                #--Adding correct <IC>_SPOSecondaryAdminsEmail if not found from scaAdmins
                if ($adminsUPN -notcontains $SPOSecondaryAdmins -and $_.TemplateId -notmatch 'TEAMCHANNEL'){
                    LogWrite -Message "-Adding [$SPOSecondaryAdminsEmail] ..."
                    Add-PnPSiteCollectionAdmin -Owners $SPOSecondaryAdminsEmail -Connection $psConn
                    $updateToDB = $true
                    #adding to report
                    $null = $Report.Add([PSCustomObject]@{
                        ICName = $ICName
                        URL = $psURL
                        Template = $_.TemplateId
                        Admins = $adminsEmail
                        SecondaryAdmins = $SPOSecondaryAdminsEmail
                        Action = "Add"
                    })
                }
                #--Remove oldIC_SPOSecondaryAdminsEmail if find wrong <IC>_SPOSecondaryAdminsEmail and remove svc if any
                $scaAdminsRemoved = @($spoAdmSvc,$cloudSvc,$scaAdminsNotAllowed)
                $scaAdminsRemoved = ($admins.Where({$_.LoginName.SubString($_.LoginName.LastIndexOf("|")+1) -in $scaAdminsRemoved})).LoginName

                if ($scaAdminsRemoved){
                    LogWrite -Message "-Removing [$scaAdminsRemoved] ..."
                    Remove-PnPSiteCollectionAdmin -Owners $scaAdminsRemoved -Connection $psConn
                    $updateToDB = $true
                    #adding to report
                    $null = $Report.Add([PSCustomObject]@{
                        ICName = $ICName
                        URL = $psURL
                        Template = $_.TemplateId
                        Admins = $adminsEmail
                        SecondaryAdmins = $scaAdminsRemoved 
                        Action = "Remove"                       
                    })
                }
                #--Validate if the site is updated after add/remove SCA then update to Sites table
                if ($updateToDB){
                    $site = ValidateSite -Url $psURL -SiteContext $psConn
                    if($site.Status -eq 'Active'){                
                        $scaEmails = (Get-PnPSiteCollectionAdmin -Connection $psConn).Email
                        $scaEmails = $scaEmails -join ";"
                        UpdateSCA $script:connectionString $site $scaEmails -SitesType Sites
                        LogWrite -Message "-Updated [$scaEmails] to DB."  
                    }
                }
            }
            catch{
                LogWrite -Level ERROR "-Unexpected Error [$psUrl]: $_ "             
            }
            finally{
                DisconnectPnpOnlineOAuth -Context $psConn
            }

        }}
        
    }}
    $attachedFiles = @()    
    if ($Report -ne $null -and $Report.Count -gt 0) {        
        LogWrite -Message "Export to csv and sending report email to SP Admins..."
        $logPath = "$($script:DirLog)"
        if (!(Test-Path $logPath)) { 
	        LogWrite -Message "Creating $logPath" 
            New-Item -ItemType "directory" -Path $logPath -Force
	    }
        LogWrite -Message "Generating Log files..." 
        $reportFile = "$logPath\SitesUpdatedSCA.csv"
                
        $Report | Export-Csv $reportFile -Encoding ASCII -NoTypeInformation
        $attachedFiles += $reportFile        
    }
    else {
        LogWrite -Message "There are currently no sites need to be populated SCA."        
    }
    GenerateSiteSCAReport -attachments $attachedFiles
    $script:EndTimeDailyCache = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    LogWrite -Message "[Populate SPO Site SCA] Start Time: $($script:StartTimeDailyCache)"
    LogWrite -Message "[Populate SPO Site SCA] End Time:   $($script:EndTimeDailyCache)"
    LogWrite -Message  "----------------------- [Populate SPO Site SCA] Execution Ended ------------------------"    
    #endregion    
}
Catch [Exception] {
    LogWrite -Level ERROR "-Unexpected Error: $_ "
}
Finally {
    LogWrite -Message  "----------------------- [Process SPO Sites] Completed ------------------------"
}
