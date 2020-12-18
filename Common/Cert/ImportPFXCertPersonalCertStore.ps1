<#
      ===========================================================================
      .DESCRIPTION
        Connecting to Exchange Online PowerShel V2 authenticating Using Certificate Thumbprint
        This authentication method can be considered more secure than using the local certificate with a password.
        In this method, you will need to import the certificate to the Personal certificate store. You only need to use the thumbprint to identify which certificate to use for authentication.
        Note that you only need to do this step once for the current user.
#>
## Set the certificate file path (.pfx)
<#$CertificateFilePath = 'D:\Scripting\O365DevOps\Common\Cert\SPO-GuestUsersMembershipReport.pfx'
## Get the PFX password
$mypwd = Get-Credential -UserName 'Enter password below' -Message 'Enter password below'
## Import the PFX certificate to the current user's personal certificate store.
#Import-PfxCertificate -FilePath $CertificateFilePath -CertStoreLocation Cert:\\CurrentUser\\My -Password $mypwd.Password
Import-PfxCertificate -FilePath $CertificateFilePath -CertStoreLocation Cert:\\LocalMachine\\My -Password $mypwd.Password
#>
##$pwd = $AdminP0rtal
$CertificateFilePath = 'D:\Scripting\O365DevOps\Common\Cert\SPO-Sync Operations.pfx'
## Get the PFX password
$mypwd = Get-Credential -UserName 'Enter password below' -Message 'Enter password below'
## Import the PFX certificate to the current user's personal certificate store.
#Import-PfxCertificate -FilePath $CertificateFilePath -CertStoreLocation Cert:\\CurrentUser\\My -Password $mypwd.Password
Import-PfxCertificate -FilePath $CertificateFilePath -CertStoreLocation Cert:\\LocalMachine\\My -Password $mypwd.Password
