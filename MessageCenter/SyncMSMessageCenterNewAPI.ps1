Import-Module MSAL.PS
##Declare Variables
$AppClientID = "0f6a6412-8f62-4b7a-a929-664174baf961"
$AppTenantID = "14b77578-9773-42d5-8507-251ca2dc2b06"
$AppClientSecret = ConvertTo-SecureString "=AyG.@/lHc8pZSrEk8xaaec53Ha@is=8" -AsPlainText -Force
$AppCertificatePath =  "cert:\LocalMachine\my\7BA7CBA81EDC57BF8446C549294148FB8490AD5B"
##Import Certificate
$AppCertificate = Get-Item $AppCertificatePath

$clientID = "0f6a6412-8f62-4b7a-a929-664174baf961"
$clientSecret = "=AyG.@/lHc8pZSrEk8xaaec53Ha@is=8"
$tenantDomain = "nih.onmicrosoft.com"

# production sharepoint app principal
$spAppId = "e039c3eb-19ad-44d2-9ed0-c0612fc61e77"
$spAppSecret = "Z7YSVs8wrQO338ui4l1NVX2OoZY3HWDY4pa5We2DGlI="
$Thumbprint = "7BA7CBA81EDC57BF8446C549294148FB8490AD5B"

$loginURL = "https://login.microsoftonline.com/"
$resource = "https://manage.office.com"
$logFile = "D:\Scripting\O365DevOps\MessageCenter\Logs\SyncMessageCenter"



function Write-Log {

    [CmdletBinding()]
    
    Param (    
    [Parameter(Mandatory=$False)]
    [ValidateSet("INFO", "WARNING", "ERROR", "CRITICAL")]
    $Level = "INFO",
    
    [Parameter(Mandatory=$True)]
    $Message = $null,
    
    [Parameter(Mandatory=$False)]
    $FilePath = $logFile + "\$((Get-Date -Format "MM-yyyy").Replace(':', '')).log")
    
    if (-not(Test-Path $FilePath))
    {
         New-Item $FilePath -Type file
    } 
    else 
    {    
        $DateTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $InputLine = "$DateTime $Level $Message"
        Add-Content -Path $FilePath -Value $InputLine
    }     
}

function Get-AccessToken()
{
    #$body = @{grant_type="client_credentials";resource=$resource;client_id=$clientID;client_secret=$clientSecret}
    try {
        $oauth = Get-MsalToken -ClientId $AppClientID -TenantId $AppTenantID -ClientSecret $AppClientSecret -Verbose #Invoke-RestMethod -Method Post -Uri $loginURL/$tenantDomain/oauth2/token?api-version=1.0 -Body $body
    }
    catch {        
        Write-Log -Level ERROR -Message "Get-AccessToken: $($_.Exception)"
        throw
    }    
    $headerParams = @{Authorization = "Bearer $($oauth.AccessToken)" }
    Write-Log -Level INFO -Message "Get-AccessToken: Authentication Successfull"
    return $headerParams
}

function Get-MessageCenterMessages() 
{
    $headers = Get-AccessToken
    $url = "https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/messages"  #"$resource/api/v1.0/$tenantDomain/ServiceComms/Messages" 
    Write-Log -Level INFO -Message "Get-MessageCenterMessages: Fetching Message Center Messages from Office 365 Service Communications API"
    
    $retryCount = 5
    $retryAttempts = 0
    $backOffInterval = 2

    while ($retryAttempts -le $retryCount) {
        try {
            $messageCenterMessages = Invoke-RestMethod -Method GET -Headers $headers -Uri $url
            $retryAttempts = $retryCount + 1
            Write-Log -Level INFO -Message "Get-MessageCenterMessages: Result Fetch Successful!!"

            $NextLink = $messageCenterMessages.'@odata.nextLink'
            Write-Output "Next data link $NextLink"
            
            $messages = ($messageCenterMessages.value) # | ? {$_.MessageType -eq "MessageCenter"})

            While ($NextLink -ne $Null) {

                $SecurityHealthIssuesRequest = Invoke-RestMethod -Headers $headers -Uri  $NextLink -Method GET

                $NextLink = $SecurityHealthIssuesRequest.'@odata.nextLink'

                Write-Output "Next data link $NextLink"

                $messages += ($SecurityHealthIssuesRequest.value)

            }

            return $messages
                
        }
        catch {            
            if ($retryAttempts -lt $retryCount) {
                $retryAttempts = $retryAttempts + 1        
                Write-Log -Level INFO -Message "Retry attempt number: $retryAttempts. Sleeping for $backOffInterval seconds..."
                Start-Sleep $backOffInterval
                $backOffInterval = $backOffInterval * 2
            }
            else {
                Write-Log -Level INFO -Message "Unable to getting Message Center Messages after $retryCount times."
                Write-Log -Level ERROR -Message "Get-MessageCenterMessages $($_.Exception)"
            }
        }
    }
    <#
    try {
        #$list = Invoke-RestMethod -Method GET -Headers $headers -Uri "$resource/api/v1.0/14b77578-9773-42d5-8507-251ca2dc2b06/activity/feed/subscriptions/list"         
        $messageCenterMessages = Invoke-RestMethod -Method GET -Headers $headers -Uri $url 
    }
    catch {        
        Write-Log -Level ERROR -Message "Get-MessageCenterMessages $($_.Exception)"
        throw
    } 
      
    Write-Log -Level INFO -Message "Get-MessageCenterMessages: Result Fetch Successful!!"
    $messages = ($messageCenterMessages.value | ? {$_.MessageType -eq "MessageCenter"})
    return $messages
    #>
}


Function Send_Email
{
	param($toEmail,$subject,$body,$attachements,$credentials,$isCloudMail)
    if($isCloudMail)
    {
        if($credentials -eq $null)
        {
            Write-Log -Level ERROR -Message "Send_Email: Credentials are needed to send cloud email"
        }
        else
        {
            $FromEmailAdd = "spoadmrpt@nih.onmicrosoft.com"
	        $SmtpServerList = 'smtp.office365.com'
	        $msg = new-object Net.Mail.MailMessage
	        $smtp = new-object Net.Mail.SmtpClient($SmtpServerList)
	        $msg.From = "$FromEmailAdd"
            $smtp.Credentials = $credentials
            $smtp.Port = 587
	        $smtp.EnableSsl = $true 
	        $toAddresses = $toEmail.split(';')
        }
    }
    else
    {
	    $FromEmailAdd = "citspmail-noreply@nih.gov"
	    $SmtpServerList = 'mailfwd.nih.gov'
	    $msg = new-object Net.Mail.MailMessage
	    $smtp = new-object Net.Mail.SmtpClient($SmtpServerList)
	    $msg.From = "$FromEmailAdd"
    }
	 
	$toAddresses = $toEmail.split(';')
	foreach($toAddress in $toAddresses)
	{
		$msg.To.Add($toAddress)
	}
	$msg.Subject = $subject
	$msg.Body = $body
	$msg.IsBodyHtml = $true
	foreach($attachment in $attachements)
	{
		$att = new-object Net.Mail.Attachment($attachment)
		$msg.Attachments.Add($att)
	}
	$smtp.Send($msg)
	$msg.Dispose()
}

try
{    
    Write-Log -Level INFO -Message "Script execution started"
    <#$username = "spoadmsvc@nih.gov"    
    $pswPath = "D:\Scripting\O365\Config\O365TenantPwd"   
       
    $password = Get-Content $pswPath | ConvertTo-SecureString    
    $credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $password    #>
    #$credentials = Get-Credential
       
    $messages = Get-MessageCenterMessages    
    
    # Change to prod when ready
    #$SiteUrl = "https://nih.sharepoint.com"   
    #$mcListId = "35c8691e-10f9-4bd9-9850-78d95e4266d9" # "Microsoft Message Center"
    $mcListId = "C854D594-DBFB-4024-B3E2-FC7C39C93B2D"
    #$mcListId = "C854D594-DBFB-4024-B3E2-FC7C39C93B2D"

    #Connect-PnPOnline -Url $SiteUrl -Credentials $credentials
    #Connect-PnPOnline -AppId $spAppId -AppSecret $spAppSecret -Url $SiteUrl
    $TenantName = "nih.onmicrosoft.com"
    $AppId = "9624e216-9e73-4513-9251-4d4382950420"
    $SPORootSiteURL = "https://nih.sharepoint.com/sites/GRP-CIT-O365-Admins/"
    $Thumbprint = "1C9696EB9152228A42DAEB5C7075699795311662"

    #Connect-PnPOnline -Tenant $TenantName -ClientId $AppId -Thumbprint $Thumbprint -Url $SiteUrl
    $siteContext = Connect-PnPOnline -Url $SPORootSiteURL -ReturnConnection -verbose -ClientId $AppId -Thumbprint $Thumbprint -Tenant $TenantName #-ClientSecret $spAppSecret  # -Thumbprint $Thumbprint -Tenant $TenantName # 
          
    # Get List
    $list = Get-PnPList -Identity $mcListId
    $items = Get-PnPListItem -List $list

    $text = "<style>
    table {
        border-collapse: collapse;
        width: 100%;
    }

    th, td {
        text-align: left;
        padding: 8px;
        border: 1px solid #dddddd;
    }

    th {
        background-color: #4CAF50;
        color: white;
    }
    </style>"
    $sendEmail = $false
    $text += "<p>Please review the latest messages from O365 Message Center.</p>"
    $text += "<table>"  
    $count = 1
    foreach($message in $messages)
    {   
        Write-Output "Processing $count of $($messages.Count)"
        $count = $count + 1
        $lastUpdatedDate = $null;
        if($message.lastModifiedDateTime -ne $null)
        {
            $lastUpdatedDate = ([DateTime]$message.lastModifiedDateTime).ToUniversalTime();
            $lastUpdatedDate = $lastUpdatedDate.AddSeconds(-$lastUpdatedDate.Second).AddMilliseconds(-$lastUpdatedDate.Millisecond)
        } 
	    $existingItems = @($items | ? {$_["MessageId"] -eq $message.id -and $_["LastUpdatedTime"] -eq $lastUpdatedDate})
        #$existingItems = @($items | ? {$_["MessageId"] -eq $message.Id})      
        $NewListItem = $null      
        if(($existingItems -eq $null -or $existingItems.Count -eq 0) -and $message.id -ne $null -and $message.id.length -gt 2)
        {  
            $message.id
            $message.title    
            $message.id.length        
            $sendEmail = $true
            $startDate = $null;
            if($message.startDateTime -ne $null)
            {
                $startDate = ([DateTime]$message.startDateTime).ToUniversalTime();
            } 
            $endDate = $null;
            if($message.endDateTime -ne $null)
            {
                $endDate = ([DateTime]$message.endDateTime).ToUniversalTime();
            } 
            $actionRequiredbyDate = $null;
            if($message.actionRequiredByDateTime -ne $null)
            {
                $actionRequiredbyDate = ([DateTime]$message.actionRequiredByDateTime).ToUniversalTime();
            }             
            $externalLink = '';
            iF($message.details -ne $null){
                if($message.details[0].name -eq 'BlogLink')
                {
                    $externalLink = $message.details[0].value
                } 
                elseif($message.details[0].name -eq 'ExternalLink')
                {    
                    $externalLink = $message.details[1].value
                } 
                elseif($message.details[1].name -eq 'ExternalLink')
                {    
                    $externalLink = $message.details[1].value
                }    
            }
            Write-Output "External Link $externalLink"
	        $NewListItem = @{
                                MessageId=$message.id;
                                Title=$message.title;
                                Category=$message.category;
                                Messages=$message.body.content;
                                #ActionType=$message.ActionType;
                                #ActionRequiredByDate=$actionRequiredbyDate;
                                ActBy=$actionRequiredbyDate;
                                StartTime=$startDate;
                                EndTime=$endDate;
                                LastUpdatedTime=$lastUpdatedDate;
                                #Status=$message.Status;
                                Tags=$message.tags;
                                Message_x0020_Category=$message.services;
                                UrgencyLevel=$message.severity;                                
                                ExternalLink=$externalLink;                                
                                CITActionStatus="New";
                            }    
            $text += "<tr><td>Message Id</td><td><b>$($NewListItem.MessageId)</b></td></tr>"            
            $text += "<tr><td>Title</td><td>$($NewListItem.Title)</td></tr>"	            
            $text += "<tr><td>Category</td><td>$($NewListItem.Category)</td></tr>"	                            	        
            $text += "<tr><td>Message</td><td><p>$($NewListItem.Messages)</p></td></tr>"       	        
            $text += "<tr><td>Services</td><td>$($NewListItem.Message_x0020_Category)</td></tr>"            
	        $text += "<tr><td>Action Required By</td><td>$($NewListItem.ActBy)</td></tr>"	            
            $text += "<tr><td>Published On</td><td>$($NewListItem.StartTime)</td></tr>"	       
            $text += "<tr><td>Expires On</td><td>$($NewListItem.EndTime)</td></tr>"	       	        
            $text += "<tr><td>Last Updated On</td><td>$($NewListItem.LastUpdatedTime)</td></tr>"            	        
	        $text += "<tr><td>Tags</td><td>$($NewListItem.Tags)</td></tr>"	        
            $text += "<tr><td>Urgency Level</td><td>$($NewListItem.UrgencyLevel)</td></tr>"
            $text += "<tr><td>More Information</td><td>$($NewListItem.ExternalLink)</td></tr>"
            $text += "<tr><td colspan='2'>&nbsp;</td></tr>"	
            Add-PnPListItem -List $list -Values $NewListItem            
        } 
        else
        {
            # Workaround code
            <#foreach($existingItem in $existingItems)
            {
                $properties = @{                               
                                LastUpdatedTime=$lastUpdatedDate;                                
                            } 
                $existingItem["LastUpdatedTime"]
                $lastUpdatedDate
                $retval = Set-PnPListItem -Identity $existingItem.Id -List $list -Values $properties
            }#>
        }       
    }
    $text += "</table>"
    if($sendEmail -eq $false)
    {
        $text = "<P>There are currently no new messages in Message Center</p>"
    }
    $subject = "[MC-NIH-GCC] Daily Updates"
    $toAddress = "CITM365CollabTenantAdmins@mail.nih.gov;cesops@nih.gov;GRP-CIT-O365-Admins@groups.nih.gov;citdcsdesktopengteam@mail.nih.gov" 
    #$toAddress = "cithsssharepointhosting@mail.nih.gov" 
    #$toAddress = "rahul.babar@nih.gov" 
    Send_Email -toEmail $toAddress -subject $subject -body $text 
}
catch
{
    Write-Host $_.Exception
    Write-Log -Level ERROR -Message $_.Exception
}
finally
{
    Disconnect-PnPOnline -Connection $siteContext
    Write-Log -Level INFO -Message "Script execution completed"
}

