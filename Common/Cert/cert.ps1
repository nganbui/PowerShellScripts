<#$certPath = "D:\Scripting\O365DevOps\Cert"
$cert = New-PnPAzureCertificate -CommonName "NIHSPODEV-AzureApp" -OutPfx "$($certPath)\citspdev.pfx" -OutCert "$($certPath)\citspdev.cer" -ValidYears 1
$cert
#installs the certificate to Local Machine.
Import-PfxCertificate -Exportable -CertStoreLocation Cert:\LocalMachine\My -FilePath "$($certPath)\citspdev.pfx"

#certutil -store my#>

$certPath = "D:\Scripting\O365DevOps\Common\Cert"
$cert = New-PnPAzureCertificate -CommonName "NIHSPODEV-M365Operation" -OutPfx "$($certPath)\M365Operation.pfx" -OutCert "$($certPath)\M365Operation.cer" -ValidYears 1
$cert
#installs the certificate to Local Machine.
#Import-PfxCertificate -Exportable -CertStoreLocation Cert:\LocalMachine\My -FilePath "$($certPath)\M365Operation.pfx"
