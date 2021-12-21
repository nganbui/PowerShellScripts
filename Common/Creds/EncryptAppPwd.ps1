Function Get-Root
{
    $root = ""
    try 
    {
        # Get info from PowerShell variable...
        $root = Get-Variable -Name PSScriptRoot -ValueOnly -ErrorAction Stop
    } 
    catch
    {
        $root = Split-Path $Script:MyInvocation.MyCommand.Path
    }
    return $root
} 

$pwdPath = Get-Root
$o365AppPwdFile = "$($pwdPath)\O365AppPwd"

read-host "Please enter O365 App Secret" -assecurestring | convertfrom-securestring | out-file $o365AppPwdFile
