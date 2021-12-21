#Import-Certificate -FilePath "D:\Scripting\O365DevOps\Common\Cert\SPO-M365OperationSupport.cer" -CertStoreLocation Cert:\LocalMachine\My
#Import-PfxCertificate -FilePath D:\Scripting\O365DevOps\Common\Cert\SPO-M365.pfx -CertStoreLocation Cert:\LocalMachine\My -Exportable
#Get-ChildItem Cert:\LocalMachine\My\1C9696EB9152228A42DAEB5C7075699795311662 | Remove-Item

# Import PfxCert - Run on target server
# Where the PfxCert in
$CerOutputPath     = "D:\Scripting\O365DevOps\Common\Cert\SPO-M365OperationSupport.pfx"
$StoreLocation     = "Cert:\LocalMachine\My"
# import cert
Get-ChildItem -Path $CerOutputPath | Import-PfxCertificate -CertStoreLocation $StoreLocation -Exportable

#Get-ChildItem -Path cert:\localMachine\my\1C9696EB9152228A42DAEB5C7075699795311662 | Export-PfxCertificate -FilePath \\nihspodev\d$\Scripting\O365DevOps\Common\Cert\SPO-M365OperationSupport.pfx -ProtectTo nih\aabuint
#Import cert to another server
#Get-ChildItem -Path D:\Scripting\O365DevOps\Common\Cert\SPO-M365OperationSupport.pfx | Import-PfxCertificate -CertStoreLocation Cert:\LocalMachine\My -Exportable

