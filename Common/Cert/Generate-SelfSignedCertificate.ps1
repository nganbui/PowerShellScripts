#certutil -store my
#cd 'D:\Scripting\O365DevOps\Common\Cert' 
#.\Create-SelfSignedCertificate.ps1 -CommonName "SPO-GuestUsersMembershipReport" -StartDate 2020-09-21 -EndDate 2022-09-21

#$pwd = $AdminP0rtal
cd 'D:\Scripting\O365DevOps\Common\Cert' 
.\Create-SelfSignedCertificate.ps1 -CommonName "SPO-Sync Operations" -StartDate 2020-12-01 -EndDate 2022-12-01