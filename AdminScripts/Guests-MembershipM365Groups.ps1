# Find guest accounts
#Connect to AzureAD and Exchange Online
#Connect-AzureAD
#Connect-ExchangeOnline
<#
$GuestUsers = Get-AzureADUser -All $true -Filter "UserType eq 'Guest'" 
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Guest in $GuestUsers) {
   $AADAccountAge = ($Guest.RefreshTokensValidFromDateTime | New-TimeSpan).Days   
    Write-Host "Processing" $Guest.DisplayName
    $i = 0; 
    $GroupNames = $Null
    # Find what Office 365 Groups the guest belongs to... if any
    $DN = (Get-Recipient -Identity $Guest.UserPrincipalName).DistinguishedName 
    $GuestGroups = (Get-Recipient -Filter "Members -eq '$Dn'" -RecipientTypeDetails GroupMailbox | Select DisplayName, ExternalDirectoryObjectId)
    If ($GuestGroups -ne $Null) {
        ForEach ($G in $GuestGroups) { 
        If ($i -eq 0) { $GroupNames = $G.DisplayName; $i++ }
        Else 
        {$GroupNames = $GroupNames + "; " + $G.DisplayName }
    }}
    $ReportLine = [PSCustomObject]@{
        UPN     = $Guest.UserPrincipalName
        Name    = $Guest.DisplayName
        Age     = $AADAccountAge
        Created = $Guest.RefreshTokensValidFromDateTime  
        Groups  = $GroupNames
        DN      = $DN}      
    $Report.Add($ReportLine) }
$Report | Sort Name | Export-CSV -NoTypeInformation D:\Scripting\O365DevOps\Common\Data\Other\GuestAccounts.csv

#>

<#
$EndDate = (Get-Date).AddDays(1); 
$StartDate = (Get-Date).AddDays(-90); 
$NewGuests = 0
$Records = (Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations "SharingInvitationCreated" -ResultSize 2000 -Formatted)
If ($Records.Count -eq 0) {
   Write-Host "No Sharing Invitations records found." }
 Else {
   Write-Host "Processing" $Records.Count "audit records..."
   $Report = @()
   ForEach ($Rec in $Records) {
      $AuditData = ConvertFrom-Json $Rec.Auditdata
      # Only process the additions of guest users to groups
      If ($AuditData.TargetUserOrGroupName -Like "*#EXT#*") {
         $TimeStamp = Get-Date $Rec.CreationDate -format g
         # Try and find the timestamp when the invitation for the Guest user account was accepted from AAD object
         Try {$AADCheck = (Get-Date(Get-AzureADUser -ObjectId $AuditData.TargetUserOrGroupName).RefreshTokensValidFromDateTime -format g) }
           Catch {Write-Host "Azure Active Directory record for" $AuditData.UserId "no longer exists" }
          If ($TimeStamp -eq $AADCheck) { # It's a new record, so let's write it out 
            $NewGuests++
            $ReportLine = [PSCustomObject][Ordered]@{
              TimeStamp    = $TimeStamp
              InvitingUser = $AuditData.UserId
              Action       = $AuditData.Operation
              URL          = $AuditData.ObjectId
              Site         = $AuditData.SiteUrl
              Document     = $AuditData.SourceFileName
              Guest        = $AuditData.TargetUserOrGroupName }      
           $Report += $ReportLine }}
      }}
$Report | Format-Table TimeStamp, Guest, Document -AutoSize
#>

$EndDate = (Get-Date).AddDays(1); $StartDate = (Get-Date).AddDays(-90); $NewGuests = 0
$Records = (Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations "Add Member to Group" -ResultSize 2000 -Formatted)
If ($Records.Count -eq 0) {
   Write-Host "No Group Add Member records found." }
 Else {
   Write-Host "Processing" $Records.Count "audit records..."
   $Report = [System.Collections.Generic.List[Object]]::new()
   ForEach ($Rec in $Records) {
      $AuditData = ConvertFrom-Json $Rec.Auditdata
      # Only process the additions of guest users to groups
      If ($AuditData.ObjectId -Like "*#EXT#*") {
         $TimeStamp = Get-Date $Rec.CreationDate -format g
         # Try and find the timestamp when the Guest account was created in AAD
         Try {$AADCheck = (Get-Date(Get-AzureADUser -ObjectId $AuditData.ObjectId).RefreshTokensValidFromDateTime -format g) }
           Catch {Write-Host "Azure Active Directory record for" $AuditData.ObjectId "no longer exists" }
         If ($TimeStamp -eq $AADCheck) { # It's a new record, so let's write it out
            $NewGuests++
            $ReportLine = [PSCustomObject]@{
              TimeStamp   = $TimeStamp
              User        = $AuditData.UserId
              Action      = $AuditData.Operation
              GroupName   = $AuditData.modifiedproperties.newvalue[1]
              Guest       = $AuditData.ObjectId }      
           $Report.Add($ReportLine) }}
      }}
Write-Host $NewGuests "new guest records found..."
$Report | Sort GroupName, Timestamp | Get-Unique -AsString | Format-Table Timestamp, Groupname, Guest