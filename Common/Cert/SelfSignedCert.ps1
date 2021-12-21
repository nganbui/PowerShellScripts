# Your tenant name (can something more descriptive as well)
$TenantName        = "citspdev.onmicrosoft.com"

# Where to export the certificate without the private key
$CerOutputPath     = "D:\Scripting\O365DevOps\Common\Cert\NihspodevAzureApp.cer"

# What cert store you want it to be in
$StoreLocation     = "Cert:\LocalMachine\My"

# Expiration date of the new certificate
$ExpirationDate    = (Get-Date).AddYears(2)


# Splat for readability
$CreateCertificateSplat = @{
    FriendlyName      = "NIHSPODEV-M365Operation"
    DnsName           = $TenantName
    CertStoreLocation = $StoreLocation
    NotAfter          = $ExpirationDate
    KeyExportPolicy   = "Exportable"
    KeySpec           = "Signature"
    Provider          = "Microsoft Enhanced RSA and AES Cryptographic Provider"       
    HashAlgorithm     = "SHA256"
}

# Create certificate
$Certificate = New-SelfSignedCertificate @CreateCertificateSplat
#$Certificate = New-SelfSignedCertificate -DnsName "IssuedToName" -CertStoreLocation "cert:\CurrentUser\My" -KeySpec Signature

# Get certificate path
$CertificatePath = Join-Path -Path $StoreLocation -ChildPath $Certificate.Thumbprint

# Export certificate without private key
Export-Certificate -Cert $CertificatePath -FilePath $CerOutputPath  | Out-Null
