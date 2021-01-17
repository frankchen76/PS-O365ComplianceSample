$credName = "SPO-M365x725618"
$cred = Get-PnPStoredCredential -Name $credName -Type PSCredential

Connect-IPPSSession -Credential $cred
Disconnect-ExchangeOnline -Confirm:$false

Get-RetentionCompliancePolicy 
$p = Get-RetentionCompliancePolicy -Identity "PSEmailRetention" | fl
$p.SharePointLocation
$policy = Get-RetentionCompliancePolicy -Identity "PSEmailRetention01" -DistributionDetail | fl
$policy.ExchangeLocation.Count

Remove-RetentionCompliancePolicy -Identity "PSEmailRetention" -ForceDeletion
Remove-RetentionCompliancePolicy -Identity "PSEmailRetention01" -ForceDeletion
Remove-RetentionCompliancePolicy -Identity "PSEmailRetention4" -Confirm:$false -ForceDeletion
Remove-RetentionCompliancePolicy -Identity "PSEmailRetention3" -ForceDeletion
Remove-RetentionCompliancePolicy -Identity "PSEmailRetention2" -ForceDeletion
Remove-RetentionCompliancePolicy -Identity "PSEmailRetention1" -ForceDeletion

for($index=1;$index -le 3;$index++){
    $name = [System.String]::Format("PSEmailRetention{0}",$index)
    Remove-RetentionCompliancePolicy -Identity $name -Confirm:$false
    Remove-RetentionCompliancePolicy -Identity $name -ForceDeletion -Confirm:$false
    Write-Host "$($name) was removed."
}

$users = "AllanD@M365x725618.OnMicrosoft.com", "ChristieC@M365x725618.OnMicrosoft.com"
$users = @("AllanD@M365x725618.OnMicrosoft.com", "ChristieC@M365x725618.OnMicrosoft.com")
New-RetentionCompliancePolicy -Name "PSEmailRetention01" -ExchangeLocation $users -SharePointLocation All -OneDriveLocation All -ModernGroupLocation All
New-RetentionComplianceRule -Name 7Years -Policy "PSEmailRetention01" -RetentionDuration 2920
Get-RetentionComplianceRule -Identity 7Years | fl



$existRetention = Get-RetentionCompliancePolicy -Identity "PSEmailRetention"
$existRetention.ExchangeLocation.GetType()

$PolicyNamePrefix = "PSEmailRetention"
Get-RetentionCompliancePolicy -DistributionDetail | Where-Object { $_.Name.StartsWith($PolicyNamePrefix) } | ft

Connect-AzureAD -Credential $cred
Get-AzureADUser -Filter "createdDateTime ge datetime'2017-07-01T08:00'" | fl
Get-AzureADUser -Filter "createdDateTime ge datetime'2017-07-01'"
$user = Get-AzureADUser -Filter "userPrincipalName eq 'frank@m365x725618.onmicrosoft.com'" | fl
# $user.ExtensionProperty
# $user | fl
# https://graph.microsoft.com/v1.0/me/events?$filter=start/dateTime ge '2017-07-01T08:00'

$users = @()
for ($i = 0; $i -lt 10; $i++) {
    $users += $i.toString()
}
$userLimit = 3
$loopCount = [int]($users.length / $userLimit) + 1
for ($index = 0; $index -lt $loopCount; $index++) {
    $ps = $userLimit * $index
    $pe = $userLimit * ($index + 1) - 1
    if ($pe -gt $users.length) {
        $pe = $users.length - 1
    }

    $processUsers = $users[$ps..$pe]
    $processUsers -join ","
}



write-host "$($users.length) users were added."
$users = $users[0..1]
$users = @()
$users.length

$ps = Get-RetentionCompliancePolicy -DistributionDetail | Where-Object { $_.Name.StartsWith("Pe") }
