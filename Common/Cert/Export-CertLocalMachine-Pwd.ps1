#certutil -store my#>
#certutil -store root#>

# Export PfxCert from existing cert - run on source server
$Thumbprint        = "7BA7CBA81EDC57BF8446C549294148FB8490AD5B"
#$pwd = "$AdminP0rtal"

$secret = Read-Host -AsSecureString
#$secret

$Encrypted = ConvertFrom-SecureString -SecureString $secret
#$Encrypted

[System.Security.SecureString] $pwd = ConvertTo-SecureString -String $Encrypted


#Get-ChildItem -Path $StoreLocation\$Thumbprint | Export-PfxCertificate -FilePath $FilePath -ProtectTo $ProtectTo

$certStoreLocation = "cert:\LocalMachine\My"
$certificatePath = $certStoreLocation + '\' + $Thumbprint
$folderPath = "D:\Download" # Where do you want the files to get saved to? The folder needs to exist.
$fileName = "SPO-Sync Operations" # What do you want to call the cert files? without the file extension
$filePath = $folderPath + '\' + $fileName
Export-Certificate -Cert $certificatePath -FilePath ($filePath + '.cer')
Export-PfxCertificate -Cert $certificatePath -FilePath ($filePath + '.pfx') -Password $pwd