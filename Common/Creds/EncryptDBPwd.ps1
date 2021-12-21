$pwdPath = "D:\Scripting\O365DevOps\Common\Creds"
$o365DBPwdFile = "$($pwdPath)\O365DBPwd"

read-host "Please enter DB Password" -assecurestring | convertfrom-securestring | out-file $o365DBPwdFile