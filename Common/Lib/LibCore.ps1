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
		$DataSet | Microsoft.PowerShell.Utility\Export-Csv $FileName -NoTypeInformation
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
			$todaysDate = Get-Date -Format "MM-dd-yyyy"
			#$fileName = [DateTime]::Now.ToString("yyyyMMdd-HHmmss")
			$Path += "\" + $todaysDate + "\SyncLog.txt"
            
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
	#tenant
	$o365TenantPwdFile = "$($profileData.Path["Cred"])\$($profileData.TenantConfig["O365TenantPwdFile"])"   
	$script:o365AdminCredential = GetCredential -AccountName $profileData.TenantConfig["O365TenantAdmin"] -AccountNameKey $o365TenantPwdFile 
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
	$script:LogFile = "$($script:RootDir)\Logs"
	if ($null -ne $logFileName) {
		$script:LogFile += "\" + $logFileName
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
	$script:TenantAdmin = "$($profileData.TenantConfig["O365TenantAdmin"])"   
	#$TenantPwdFile = "$($profileData.Path["Cred"])\$($profileData.TenantConfig["O365TenantPwdFile"])"
	#$script:o365AdminCredential = GetCredential -AccountName $script:TenantAdmin -AccountNameKey $TenantPwdFile     	
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
Function Set-AzureAppVars {
	param(
		[parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()] 		
		[ValidateScript( { Test-Path -Path $_ -PathType Leaf })]
		[string]$profilePath = $global:ProfilePath
	)	
	#--Retrieve .psd1 
	$profileData = GetPsd1Data $profilePath
    #--Azure App - [SPO-M365 Operations Support] - app for M365 Operation support
    $script:appIdOperationSupport = "$($profileData.AppConfigOperationSupport["AppId"])"
    $script:appThumbprintOperationSupport = "$($profileData.AppConfigOperationSupport["Thumbprint"])"
    $script:appCertOperationSupport = "$($profileData.AppConfigOperationSupport["CertName"])"
    #--Azure App - [SPO-GuestUsersMembershipReport] - app for Guests Activity
    $script:appIdEXOV2 = "$($profileData.AppConfigEXOV2["AppId"])"
    $script:appThumbprintEXOV2 = "$($profileData.AppConfigEXOV2["Thumbprint"])"
    #--Azure App - [SPO-AdminPortal Operations] - app for sync job
    $script:appIdAdminPortalOperation = "$($profileData.AppConfigAdminPortalOperation["AppId"])"
    $script:appThumbprintAdminPortalOperation = "$($profileData.AppConfigAdminPortalOperation["Thumbprint"])"
    $script:appCertAdminPortalOperation = "$($profileData.AppConfigAdminPortalOperation["CertName"])"
    #--Azure App - [SPO - Message Center and Health Reader] - app for message center
    $script:appIdMC = "$($profileData.AppConfigMC["AppId"])"
    $script:appSecretMC = "$($profileData.AppConfigMC["AppSecret"])"
    $script:appResourceMC = "$($profileData.AppConfigMC["Resource"])"
    #--Azure App - [SPO-M365 Reports]
    <#$script:appIdUsageReport = "$($profileData.AppConfigUsageReport["AppId"])"
    $script:appThumbprintUsageReport = "$($profileData.AppConfigUsageReport["Thumbprint"])"
    #>
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
	$script:Admin = "$($profileData.EmailConfig["Admin"])"
	$script:Support = "$($profileData.EmailConfig["Support"])"	
    
}
Function Set-StatusVars {    
	$script:Submitted = "5A6C2888-1D7F-4FFC-94FA-0A92640C7076"
    $script:Completed = "E950207B-AAF3-4B65-9BB7-689A6B6AE83D"    
    $script:InProgress = "6B934E0F-8784-461B-80D0-A4660F6D1A4E"
    $script:Pending = "6C4B8971-FAFF-4B26-BB6F-FD4C5CEA66AA"    
}
Function Set-ChangeTypeVars {    
	$script:ExternalSharing = "FF07775A-9883-4C42-89C1-947F347AE712"
    $script:TeamsDisplayName = "B252D6A3-D11A-4AB3-84A9-9DA4769DA3F2"     
}


#endregion

#region Misc
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
		[parameter(Mandatory = $false)]$Attachements    
	)
	Set-EmailVars        
	$msg = new-object Net.Mail.MailMessage
	$smtp = new-object Net.Mail.SmtpClient($script:SmtpServer)

	#From-Address
	$msg.From = $script:NoReply
	#To-Addess
	$toAddresses,
	$cc = $null
	
	#To-Address
	if (($script:Admin -ne "") -and ($script:Admin -ne $null)) {
		$toAddresses = $script:Admin -split ';'

		foreach ($addr in $toAddresses) {
			$msg.To.Add($addr)
		}
	}
    
	#CC-Address
	if (($script:Support -ne "") -and ($script:Support -ne $null)) {
		foreach ($cc in $script:Support) {
			$msg.CC.Add($cc)
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

