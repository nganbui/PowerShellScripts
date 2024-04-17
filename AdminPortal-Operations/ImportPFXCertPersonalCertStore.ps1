## Get the PFX password
$mypwd = Get-Credential -UserName 'Enter password below' -Message 'Enter password below'
## Import the PFX certificate to the current user's personal certificate store.
Import-PfxCertificate -FilePath $CertificateFilePath -CertStoreLocation Cert:\\CurrentUser\\My -Password $mypwd.Password
