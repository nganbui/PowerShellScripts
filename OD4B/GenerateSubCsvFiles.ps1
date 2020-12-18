$OneDriveList = Import-Csv -LiteralPath "C:\TEMP\OneDrive\OneDrive.csv"

$icList = @() #$OneDriveList | select IC | Get-Unique

foreach($OneDrive in $OneDriveList)
{
    if($icList.Contains($OneDrive.IC) -eq $false)
    {
        $icList += $OneDrive.IC
        $icOneDrives = $OneDriveList | ? {$_.IC -eq $OneDrive.IC} 
        $icOneDrives | Export-Csv -LiteralPath "C:\TEMP\OneDrive\$($OneDrive.IC).csv" -NoTypeInformation
    }
}

