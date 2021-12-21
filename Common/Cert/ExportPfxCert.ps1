# Export PfxCert from existing cert - run on source server
$Thumbprint        = "1C9696EB9152228A42DAEB5C7075699795311662"
# What cert store you want it to be in
$StoreLocation     = "Cert:\LocalMachine\My"
# Target Path for Pfx -target server
$FilePath = "\\nihspodev\d$\Scripting\O365DevOps\Common\Cert\SPO-M365OperationSupport.pfx"
# service account which run scheduled job
$ProtectTo = "nih\aabuint"
Get-ChildItem -Path $StoreLocation\$Thumbprint | Export-PfxCertificate -FilePath $FilePath -ProtectTo $ProtectTo