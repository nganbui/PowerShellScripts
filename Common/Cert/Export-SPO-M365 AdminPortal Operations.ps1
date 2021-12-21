#certutil -store my#>
#certutil -store root#>

# Export PfxCert from existing cert - run on source server
$Thumbprint = "D0CE39CFCDBE72E92A3233A4C3BB891DC6045431"
#$pwd = "$AdminP0rtal"

$secret = Read-Host -AsSecureString
#$secret

$Encrypted = ConvertFrom-SecureString -SecureString $secret
#$Encrypted

[System.Security.SecureString] $pwd = ConvertTo-SecureString -String $Encrypted


#Get-ChildItem -Path $StoreLocation\$Thumbprint | Export-PfxCertificate -FilePath $FilePath -ProtectTo $ProtectTo

$certStoreLocation = "cert:\LocalMachine\My"
$certificatePath = $certStoreLocation + '\' + $Thumbprint
$folderPath = "D:\Scripting\O365DevOps\Common\Cert" # Where do you want the files to get saved to? The folder needs to exist.
$fileName = "SPO-M365 AdminPortal Operations" # What do you want to call the cert files? without the file extension
$filePath = $folderPath + '\' + $fileName
Export-Certificate -Cert $certificatePath -FilePath ($filePath + '.cer')
Export-PfxCertificate -Cert $certificatePath -FilePath ($filePath + '.pfx') -Password $pwd