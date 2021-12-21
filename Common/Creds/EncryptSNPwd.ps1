$pwdPath = "D:\Scripting\O365DevOps\Common\Creds"
$SNPwdFile = "$($pwdPath)\ServiceNowPwd"

read-host "Please enter SN password" -assecurestring | convertfrom-securestring | out-file $SNPwdFile
