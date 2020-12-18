#certutil -store my#>
#certutil -store root#>

# Export PfxCert from existing cert - run on source server
$Thumbprint        = "EAADC1737F06BB0104C8E6932EB814AB0CD802E6"
$ProtectTo = "nih\aabuint"

#Get-ChildItem -Path $StoreLocation\$Thumbprint | Export-PfxCertificate -FilePath $FilePath -ProtectTo $ProtectTo

$certStoreLocation = "cert:\LocalMachine\My"
$certificatePath = $certStoreLocation + '\' + $Thumbprint
$folderPath = "D:\Download" # Where do you want the files to get saved to? The folder needs to exist.
$fileName = "SPO-GuestUsersMembershipReport" # What do you want to call the cert files? without the file extension
$filePath = $folderPath + '\' + $fileName
Export-Certificate -Cert $certificatePath -FilePath ($filePath + '.cer')
Export-PfxCertificate -Cert $certificatePath -FilePath ($filePath + '.pfx') -ProtectTo $ProtectTo