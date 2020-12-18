$pwdPath = "D:\Scripting\O365DevOps\Common\Creds"
$o365AppPwdFile = "$($pwdPath)\O365AppMCPwd"

read-host "Please enter MC Azure App Password" -assecurestring | convertfrom-securestring | out-file $o365AppPwdFile