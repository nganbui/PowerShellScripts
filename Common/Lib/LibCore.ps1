$global:ProfilePath = "$($script:RootDir)\Common\Config\PROFILE.psd1"
#region Read config data .psd1 file 
Function GetPsd1Data {
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory = $true)]
		[Microsoft.PowerShell.DesiredStateConfiguration.ArgumentToConfigurationDataTransformation()]
		[hashtable] $data
	)
	return $data
}
#endregion

#region Read credential
Function GetCredential {
	[cmdletBinding()]
	Param 
	( 
		[parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AccountName,
		[parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AccountNameKey    
	)
    
	$pwdFile = Get-Content $AccountNameKey | ConvertTo-SecureString
	$credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AccountName, $pwdFile  

	return  $credential
}
#endregion

#region Create Folder/Files
Function Create-Directory {
	param($dirPath)    
	If (!(test-path $dirPath -PathType Container)) {
		New-Item -ItemType Directory -Force -Path $dirPath
	}	
}
Function Create-File {
	param($filePath, $header = "")
	if (!(Test-Path $filePath)) { 
		if ($header -eq "") {
			Out-File $filePath -Encoding ascii -force #csv file
		}
		else {
			$header | Out-File $filePath -Encoding ascii -force #csv file
		}
	}
}
#endregion

#region Export .csv/excel file
Function ExportCSV {
	param(
		[Parameter(Mandatory = $true)] $DataSet,
		[Parameter(Mandatory = $true)] $FileName
	)    
	if ($DataSet -ne $null) {        
		$DataSet | Export-Csv -LiteralPath $FileName -Force -NoTypeInformation
	}
	else {
		Write-Verbose -Message "Object is null. Couldn't save file to '$($FileName)"            
	}

}

function Export-Excel {
	[cmdletBinding()]
	Param(
		[Parameter(ValueFromPipeline=$true)]
		[string]$junk        )
	begin{
		$header = $null
		$row = 1
        if($Workbook.WorkSheets.item(1).name -eq "sheet1"){
            $Worksheet = $Workbook.WorkSheets.item(1)
        }
        else{
            $Worksheet = $Workbook.Worksheets.Add()
        }
	}
	process{
		if(!$header){
			$i = 0
			$header = $_ | Get-Member -MemberType NoteProperty | select name
			$header | %{$Worksheet.cells.item(1,++$i)=$_.Name}
		}
		$i = 0
		++$row
		foreach($field in $header){
			$Worksheet.cells.item($row,++$i)=$($_."$($field.Name)")
		}
	}
	end{
        $Worksheet.Name = $($_."Office365Report")
        $Worksheet.Columns.AutoFit() | out-null
        $Worksheet = $null
        $header = $null
	}
}    

#endregion
Function ReplaceSingleQuote {
    param($inStr)
    $outStr = $null
    if ($inStr -ne $null) {
        $outStr = $inStr.Replace("'", "''")
    }
    return $outStr
}
#region Write log file
Function LogWrite { 
	[CmdletBinding()] 
	Param 
	( 
		[Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
		[ValidateNotNullOrEmpty()] 
		[Alias("LogMessage")] 
		[string]$Message, 
 
		[Parameter(Mandatory = $false)] 
		[Alias('LogPath')] 
		[string]$Path = $script:LogFile, 
         
		[Parameter(Mandatory = $false)] 
		[ValidateSet("ERROR", "WARN", "INFO", "VERBOSE")] 
		[string]$Level = "INFO", 
         
		[Parameter(Mandatory = $false)] 
		[switch]$NoClobber
	) 
          
	# Logging process Start
	 
	Begin { 
		# Set VerbosePreference to Continue so that verbose messages are displayed. 
		$VerbosePreference = 'Continue'        
            
	} 
	Process { 
		if ($Path -ne "") {
			#$todaysDate = Get-Date -Format "MM-dd-yyyy"            
            #$FormattedDate = [DateTime]::Now.ToString("HH.mm.ss")            
			#$fileName = [DateTime]::Now.ToString("yyyyMMdd-HHmmss") + ".log"
            #$todaysDate = "$todaysDate\$FormattedDate"            
            #$Path = "$Path\$fileName"
			#$Path += "\" + $todaysDate + "\SyncLog.txt"
            
           
			# If the file already exists and NoClobber was specified, do not write to the log. 
			if ((Test-Path $Path) -AND $NoClobber) { 
				Write-Error "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name." 
				Return 
			} 
 
			# If attempting to write to a log file in a folder/path that doesn't exist create the file including the path. 
			elseif (!(Test-Path $Path)) { 
				Write-Verbose "Creating $Path." 
				$NewLogFile = New-Item $Path -Force -ItemType File 
			} 
 
			else { 
				# Nothing to see here yet. 
			} 
 
			# Format Date for our Log File 
			$FormattedDate = Get-Date -Format "MM-dd-yyyy HH:mm:ss" 
 
			# Write message to error, warning, or verbose pipeline and specify $LevelText 
			switch ($Level) { 
				'Error' { 
					Write-Error $Message 
					$LevelText = 'ERROR:' 
				} 
				'Warn' { 
					Write-Warning $Message 
					$LevelText = 'WARNING:' 
				} 
				'Info' { 
					Write-Host $Message 
					$LevelText = 'INFO:' 
				}
				'Verbose' {
					$LevelText = 'DEBUG:' 
				} 
			}

			# create a mutex, so we can lock the file while writing to it
			$mutex = New-Object System.Threading.Mutex($false, 'LogMutex')
			[void]$mutex.WaitOne()

			if ($Level -ne "Verbose") {
				"$FormattedDate $LevelText $Message" | Out-File -FilePath $Path -Append 
			}
			else {
				if ($debugLog -eq $true) {
					Write-Verbose $Message
					"$FormattedDate $LevelText $Message" | Out-File -FilePath $Path -Append 
				}
			}
			$mutex.ReleaseMutex()
				
		}
	} 
	End { 
	} 
}
#endregion

#region Global variables
Function Set-GlobalVars {
	param(
		[parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()] 
		[string]$profilePath = $global:ProfilePath 
	) 
	#Retrieve psd1 using Import-PowerShellDataFile is available in powershell version 7.0
	$profileData = GetPsd1Data $profilePath
	#URL
	$script:SPOAdminCenterURL = "$($profileData.TenantConfig["AdminCenterUrl"])"
	$script:SPORootSiteURL = "$($profileData.TenantConfig["RootSiteUrl"])"	
	#db
	$o365DBPwdFile = "$($profileData.Path["Cred"])\$($profileData.DBConfig["DBPwdFile"])"       
	$o365DBCredential = GetCredential -AccountName $profileData.DBConfig["DBUser"] -AccountNameKey $o365DBPwdFile 
	$o365DBDecryptedPwd = ($o365DBCredential.GetNetworkCredential()).Password    
	$script:ConnectionString = "Data Source=$($profileData.DBConfig["DBServer"]);Initial Catalog=$($profileData.DBConfig["DBName"]);Integrated Security=False;User ID=$($profileData.DBConfig["DBUser"]);Password=$($o365DBDecryptedPwd)" 
	#Azure App
	$o365AppKeyFile = "$($profileData.Path["Config"])\$($profileData.AppConfig["AppSecret"])"   
	$script:o365AppCredential = GetCredential -AccountName $profileData.AppConfig["AppId"] -AccountNameKey $o365AppKeyFile        
}
Function Set-LogFile {
	param(
		[parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()] 		
		[ValidateScript( { Test-Path -Path $_ -PathType Leaf })]
		[string]$profilePath = $global:ProfilePath,
		[parameter(Mandatory = $false)]
		[string]$logFileName
	)	
	#Retrieve .psd1 
	$profileData = GetPsd1Data $profilePath
	#$script:LogFile = "$($profileData.Path["Log"])"
	<#$script:LogFile = "$($script:RootDir)\Logs"
	if ($null -ne $logFileName) {
		$script:LogFile += "\" + $logFileName
	}#>
    if ($null -ne $logFileName) {
        $logFileName = "$($script:RootDir)\Logs\$logFileName"
    }
    $todaysDate = [DateTime]::Now.ToString("MM-dd-yyyy")
    $FormattedDate = [DateTime]::Now.ToString("HH.mm.ss")
    $fileName = [DateTime]::Now.ToString("yyyyMMdd-HHmmss") + ".log"
    $script:DirLog = "$logFileName\$todaysDate\$FormattedDate"
    $script:LogFile = "$($script:DirLog)\$fileName"
	
    foreach ($dirPath in $script:DirLog) {        
        Create-Directory $dirPath
    }
}

Function Set-DataFile {
	param(
		[parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()] 		
		[ValidateScript( { Test-Path -Path $_ -PathType Leaf })]
		[string]$profilePath = $global:ProfilePath
	)	
	#Retrieve .psd1 
	$profileData = GetPsd1Data $profilePath
	#$script:CacheDataPath = "$($profileData.Path["Data"])"
	$script:CacheDataPath = "$($script:RootDir)\Common\Data"
    
}
Function Set-TenantVars {
	param(
		[parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()] 		
		#[ValidateScript({  Test-Path -Path $_ -PathType Leaf  })]
		[string]$profilePath = $global:ProfilePath
	)
	#Get-Content -Path $FilePath
	#Retrieve .psd1 
	$profileData = GetPsd1Data $profilePath
	#URL
	$script:SPOAdminCenterURL = "$($profileData.TenantConfig["AdminCenterUrl"])"
	$script:SPORootSiteURL = "$($profileData.TenantConfig["RootSiteUrl"])"    
	#tenant	
    $script:TenantId = "$($profileData.TenantConfig["Id"])"
    $script:TenantName = "$($profileData.TenantConfig["Name"])"
	#$script:TenantAdmin = "$($profileData.TenantConfig["O365TenantAdmin"])"
    $script:CloudSvcForProvision = "$($profileData.TenantConfig["CloudSvcForProvision"])"   
}
Function Set-DBVars {
	param(
		[parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()] 		
		[ValidateScript( { Test-Path -Path $_ -PathType Leaf })]
		[string]$profilePath = $global:ProfilePath
	)	
	#Retrieve .psd1 
	$profileData = GetPsd1Data $profilePath
	#db+
	$o365DBPwdFile = "$($profileData.Path["Cred"])\$($profileData.DBConfig["DBPwdFile"])"       
	$o365DBCredential = GetCredential -AccountName $profileData.DBConfig["DBUser"] -AccountNameKey $o365DBPwdFile 
	$o365DBDecryptedPwd = ($o365DBCredential.GetNetworkCredential()).Password    
	$script:ConnectionString = "Data Source=$($profileData.DBConfig["DBServer"]);Initial Catalog=$($profileData.DBConfig["DBName"]);Integrated Security=False;User ID=$($profileData.DBConfig["DBUser"]);Password=$($o365DBDecryptedPwd)" 
    #$script:SQLConnectionString = "Data Source=$($profileData.DBConfig["DBServer"]);Initial Catalog=$($profileData.DBConfig["DBName"]);Integrated Security=SSPI;"
}

Function Set-SNVars {
	param(
		[parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()] 		
		[ValidateScript( { Test-Path -Path $_ -PathType Leaf })]
		[string]$profilePath = $global:ProfilePath
	)	
	#Retrieve .psd1 
	$profileData = GetPsd1Data $profilePath
	#Service Now
	$script:SNServcieUrl = "$($profileData.SNConfig["ServiceUrl"])"
    $script:SNGroup = "$($profileData.SNConfig["Group"])"
    $script:SNDefaultAssignee = "$($profileData.SNConfig["DefaultAssignee"])"    
    $SNPwdFile = "$($profileData.Path["Cred"])\$($profileData.SNConfig["PwdFile"])"
    $SNCredential = GetCredential -AccountName $profileData.SNConfig["AdminAccount"] -AccountNameKey $SNPwdFile
    $SNAdminpassword = ($SNCredential.GetNetworkCredential()).Password
    $SNAdminAccount = "$($profileData.SNConfig["AdminAccount"])"
	$script:SNAuthInfo = "<urn:userName>$SNAdminAccount</urn:userName><urn:password>$SNAdminpassword</urn:password>"
    
}

Function Set-AzureAppVars {
	param(
		[parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()] 		
		[ValidateScript( { Test-Path -Path $_ -PathType Leaf })]
		[string]$profilePath = $global:ProfilePath
	)	
	#--Retrieve .psd1 
	$profileData = GetPsd1Data $profilePath
    #--Azure App - [SPO-M365 Operations Support] - app for M365 Operation support and provisioning
    $script:appIdOperationSupport = "$($profileData.AppConfigOperationSupport["AppId"])"
    $script:appThumbprintOperationSupport = "$($profileData.AppConfigOperationSupport["Thumbprint"])"
    $script:appCertOperationSupport = "$($profileData.AppConfigOperationSupport["CertName"])"    
    #--Azure App - [SPO-Sync Operations] - app for sync job
    $script:appIdAdminPortalOperation = "$($profileData.AppConfigAdminPortalOperation["AppId"])"
    $script:appThumbprintAdminPortalOperation = "$($profileData.AppConfigAdminPortalOperation["Thumbprint"])"
    $script:appCertAdminPortalOperation = "$($profileData.AppConfigAdminPortalOperation["CertName"])"
    #--Azure App - [SPO-MC and Usage Reports] - app for message center and usage report
    $script:appIdUsageReport = "$($profileData.AppConfigUsageReport["AppId"])"
    $script:appThumbprintUsageReport = "$($profileData.AppConfigUsageReport["Thumbprint"])"
    $script:appCertUsageReport = "$($profileData.AppConfigUsageReport["CertName"])"
    $script:appResourceMC = "$($profileData.AppConfigUsageReport["Resource"])"    
    #--Azure App - [SPO-PowerBI.Workspaces] - app for PowerBI Workspaces
    $script:appIdPowerBIWorkspace = "$($profileData.AppConfigPowerBIWorkspace["AppId"])"
    $script:appThumbprintPowerBIWorkspace = "$($profileData.AppConfigPowerBIWorkspace["Thumbprint"])"
    $script:appCertPowerBIWorkspace = "$($profileData.AppConfigPowerBIWorkspace["CertName"])"
    
}

Function Set-EmailVars {
	param(
		[parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()] 		
		[ValidateScript( { Test-Path -Path $_ -PathType Leaf })]
		[string]$profilePath = $global:ProfilePath
	)	
	#Retrieve .psd1 
	$profileData = GetPsd1Data $profilePath
	#Smtp
	$script:SmtpServer = "$($profileData.EmailConfig["SmtpServer"])"
	$script:NoReply = "$($profileData.EmailConfig["NoReply"])"
    $script:From = "$($profileData.EmailConfig["From"])"
	$script:Admin = "$($profileData.EmailConfig["Admin"])"
	$script:Support = "$($profileData.EmailConfig["Support"])"	
    
}
Function Set-StatusVars {    
	$script:Submitted = "5A6C2888-1D7F-4FFC-94FA-0A92640C7076"
    $script:Cancelled = "6BFA67B8-18E2-46CB-A8B8-651F36C4C9A5"
    $script:Completed = "E950207B-AAF3-4B65-9BB7-689A6B6AE83D"    
    $script:InProgress = "6B934E0F-8784-461B-80D0-A4660F6D1A4E"
    $script:Pending = "6C4B8971-FAFF-4B26-BB6F-FD4C5CEA66AA"    
}
Function Set-SiteRequestTypeVars {    
	$script:Provision = "2FC3F2DC-08D2-4E12-AB5E-8FE68C716318"
    $script:Decomission = "1B897CAF-75EF-478F-9A6C-A6EB283CF3A6"     
}
Function Set-ChangeTypeVars {    
	$script:ExternalSharing = "FF07775A-9883-4C42-89C1-947F347AE712"
    $script:HiddenOutlookGAL = "AC2CB606-68A9-490F-9971-EFAB6FEB81E0"
    $script:PrivateChannelAccess = "EB05368F-472B-491C-AF18-2D5BE51136DD"
    $script:EnableAppCatalog = "4D2A196D-C1A7-4F26-B2EF-2E6008B50BE0"
    $script:EnableSiteCustomization = "F8372500-FB02-4BDA-B13A-6C531E2432A0"
    $script:RegisterHubSite = "A7C753F5-AA0A-411F-AC51-F3D01A1348EE"
    $script:SiteAdminAccess = "F5B0B64D-6670-4AC9-94EF-BC5CE1A52A84"
    $script:OwnerAccess = "414A5774-588A-49A1-B01D-868934329D08"
    $script:ChangeDisplayName = "B252D6A3-D11A-4AB3-84A9-9DA4769DA3F2"
    $script:ChangeDescription = "2096A9FE-8AFF-4580-975D-3283B975F9A0"
    $script:ChangePrivacy = "3B4B6991-9F62-46F6-9508-59E51E48B3CF"

    #$script:TeamsDisplayName = "B252D6A3-D11A-4AB3-84A9-9DA4769DA3F2"
    #$script:TeamsDescription = "2096A9FE-8AFF-4580-975D-3283B975F9A0"    
    #$script:TeamsPrivacy = "3B4B6991-9F62-46F6-9508-59E51E48B3CF"
    #$script:GroupDisplayName = "A0404BB8-E03E-4883-9957-0FA9ED75B2B2"
    #$script:GroupDescription = "75A0044D-872B-40BE-B862-630FD67320A9"
    #$script:GroupPrivacy = "9E5F7C7A-B857-494D-A1AC-BCBE9EF2DFD8"    
    #$script:TeamsOwnerAccess = "414A5774-588A-49A1-B01D-868934329D08"
    #$script:GroupOwnerAccess = "B4F8A98B-9D46-40B6-9FDB-0A4D1C271F12"
}
Function Set-MiscVars{
    $script:M365Group = "GROUP#0" 
    $script:MSTeams = "Team" 
    $script:TEAMCHANNEL = "TEAMCHANNEL#0" 
    $script:GuidEmpty = [system.guid]::empty
    $script:EnabledTeam = "Team"
    $script:MaxRequests = 30
    $script:ProcessNew = "[Process-NewSite]"
    $script:ProcessInProgress = "[Process-InProgressSite]"
    $script:ProcessDecommission = "[Process-Decommission]"
    $script:PowerBIServicePrincipalId = "827e23fa-1e75-4259-a3f4-9ef00e3b39f4"
    $script:PowerBIUrl = "https://app.powerbi.com"
}
#endregion

#region Misc
Function LookupM365Sku{
    [CmdletBinding()]
    param($Skus,
          [string]$Sku
        )
    if ($sku -in $Skus.Keys)
    {
        return $Skus.$sku | Get-Unique
    }
    <#else
    {
        Return ('{0} not found' -f $sku)
    }#>
}
Function ReplaceSingleQuote {
	param($inStr)
	$outStr = $null
	if ($inStr -ne $null) {
		$outStr = $inStr.Replace("'", "''")
	}
	return $outStr
}
Function SendEmail {
	[cmdletBinding()]
	Param 
	(         
		[parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Subject,
		[parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Body,
        [parameter(Mandatory = $false)][ValidateNotNullOrEmpty()][string]$To,        
        [switch]$EnabledCc,
        [parameter(Mandatory = $false)][ValidateNotNullOrEmpty()][string]$Cc,
		[parameter(Mandatory = $false)]$Attachements    
	)
	Set-EmailVars        
	$msg = new-object Net.Mail.MailMessage
	$smtp = new-object Net.Mail.SmtpClient($script:SmtpServer)    

	#From-Address	
    $msg.From = New-Object System.Net.Mail.MailAddress $script:NoReply,$script:From
    #To-Address	
    $ToAddess = $script:Admin      
    if (($To -ne "") -and ($To-ne $null)) {
		    $ToAddess = $To
	    }
    $ToAddess = @($ToAddess -split ";")
    foreach($addr in $ToAddess) {
        $msg.To.Add($addr)
        }
	
    #CC-Address
    if($EnabledCc.IsPresent){
        $CcAddress = $script:Support
        if (($Cc -ne "") -and ($Cc-ne $null)) {
                $CcAddress = $Cc
                }
        $CcAddress = @($CcAddress -split ";")
        if ($CcAddress){
            foreach ($c in $CcAddress) {
		        $msg.CC.Add($c)
	        }
        }
    }

	#Email Subject
	$msg.Subject = $Subject

	#Email Body
	$msg.Body = $Body
	$msg.IsBodyHtml = $true

	#Email Attachments
	if ($Attachements -ne $null) {
		foreach ($attachment in $Attachements) {
			$attObj = new-object Net.Mail.Attachment($attachment)
			$msg.Attachments.Add($attObj)
		}
	}

	#Send Email
	$smtp.Send($msg)

	#Dispose msg object
	$msg.Dispose()
}
#endregion

