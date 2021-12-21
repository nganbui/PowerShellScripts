Function Get-Root {
    $root = ""
    try {       
        $root = Get-Variable -Name PSScriptRoot -ValueOnly -ErrorAction Stop
    } 
    catch {
        $root = Split-Path $Script:MyInvocation.MyCommand.Path
    }
    return $root
} 

$pwdPath = Get-Root

#$o365TenantPwdFile = "$($pwdPath)\O365TenantPwd"
#$o365AppAdminPortalPwdFile = "$($pwdPath)\O365AppAdminPortal"
#$o365AppPwdFile = "$($pwdPath)\O365AppPwd"
#$o365TeamsPolicyManagementPwdFile = "$($pwdPath)\O365-App-NIHTeamsPolicyManagement"
#$o365DBPwdFile = "$($pwdPath)\O365DBPwd"
#$SNPwdFile = "$($pwdPath)\ServiceNowPwd"
#$PwdFile = "$($pwdPath)\OnPremAdmPwd"
$o365DBPwdFile = "$($pwdPath)\O365DBProdPwd"


#read-host "Please enter O365 Tenant password" -assecurestring | convertfrom-securestring | out-file $o365TenantPwdFile
#read-host "Please enter App AdminPortal password" -assecurestring | convertfrom-securestring | out-file $o365AppAdminPortalPwdFile
#read-host "Please enter Azure App Secret" -assecurestring | convertfrom-securestring | out-file $o365AppPwdFile
#read-host "Please enter DB Password" -assecurestring | convertfrom-securestring | out-file $o365DBPwdFile
#read-host "Please enter SN password" -assecurestring | convertfrom-securestring | out-file $SNPwdFile
#read-host "Please enter Password for On-Prem Admin Account:" -assecurestring | convertfrom-securestring | out-file $PwdFile

read-host "Please enter Password for On-Prem Admin Account:" -assecurestring | convertfrom-securestring | out-file $o365DBPwdFile



