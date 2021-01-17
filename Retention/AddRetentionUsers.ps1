function Get-GermanyUsers {
    [CmdletBinding()]
    param (
        [PSCredential] $cred    
    )
    $ret = @();
    $null = Connect-AzureAD -Credential $cred
    # get last 7 days created users which can be combined with officeLocation Country eq 'United State'
    Get-AzureADUser -Top 10 -Filter "createdDateTime ge datetime'2019-10-01' and UserType eq 'Member'" | ForEach-Object -Process { $ret += $_.UserPrincipalName }
    return $ret
}

function Get-AvailableRetentionPolicy {
    param (
        [string] $PolicyNamePrefix,
        [int] $AccountLimit
    )
    $ret = $null
    $policies = Get-RetentionCompliancePolicy -DistributionDetail | Where-Object { $_.Name.StartsWith($PolicyNamePrefix) }
    foreach ($p in $policies) {
        if ($p.ExchangeLocation.Count -lt $AccountLimit) {
            $ret = $p
            break
        }
    }
    return $ret
}

function Get-AvailableRetentionPolicyCount {
    param (
        [string] $PolicyNamePrefix
    )
    $ret = 0
    $policies = Get-RetentionCompliancePolicy -DistributionDetail | Where-Object { $_.Name.StartsWith($PolicyNamePrefix) }
    if ($null -ne $policies) {
        if ($policies -is [array]) {
            $ret = $policies.count
        }
        else {
            $ret = 1
        }
    }
    return $ret
}

function Add-UserToRetentionPolicy {
    param (
        [string]$PolicyName,
        [string[]]$Users
    )
    
    Set-RetentionCompliancePolicy -Identity $PolicyName -AddExchangeLocation $Users
}
function Create-RetentionPolicy {
    param (
        [string]$PolicyNamePrefix,
        [string[]]$Users
    )

    $totalPolicyCount = Get-AvailableRetentionPolicyCount -policyNamePrefix $PolicyNamePrefix
    $policyName = [System.String]::Format("{0}{1}", $policyNamePrefix, $totalPolicyCount + 1)
    $ruleName = [System.String]::Format("Rule{0}", $totalPolicyCount + 1)

    # Create Retention Policy and enable Exchange with users, SPO, OD4B and O365Group. 
    $null = New-RetentionCompliancePolicy -Name $policyName -ExchangeLocation $users -SharePointLocation All -OneDriveLocation All -ModernGroupLocation All
    write-host "New Retention Policy $($policyName) were added."

    # Create Retention Rule and set retention to 7 Year 365*7=2920
    $null = New-RetentionComplianceRule -Name $ruleName -Policy $policyName -RetentionDuration 2920
    write-host "New Retention Policy $($ruleName) were added to $($policyName)."

    $ret = Get-RetentionCompliancePolicy -identity $policyName
    return $ret
        
}

$userLimit = 1000               # the max user account for a retention policy. default is 1000. for testing, you can change it to lower value
$pnPrefix = "PSEmailRetention"  # the retention policy name prefix, the PS will add sequence number after that. 
$credName = "SPO-M365x725618"   # the credential name which PS uses to connect to EXO. You need to create this at credential store 
$cred = Get-PnPStoredCredential -Name $credName -Type PSCredential

# Retrieved Germany users
$users = Get-GermanyUsers -cred $cred
write-host "Retrieved $($users.length) users"

# Connect to EXO for Retention PS
Connect-IPPSSession -Credential $cred
write-host "Connected EXO online using credential $($cred.UserName)"

$availPolicy = Get-AvailableRetentionPolicy -PolicyNamePrefix $pnPrefix -AccountLimit $userLimit

#if ($null -eq $availPolicy) {
#    $availPolicy = Create-RetentionPolicy -PolicyNamePrefix $pnPrefix -Users $users
#}

if ($null -ne $availPolicy) {
    if ($users.length -le $userLimit - $availPolicy.ExchangeLocation.count) {
        # If user's count is less than available count
        Add-UserToRetentionPolicy -PolicyName $availPolicy.Name -Users $users
        write-host "$($users.length) users were added to $($availPolicy.Name)."
        # reset users array to empty
        $users = @()
    }
    else {
        # If user's count is more than available count, add the available users first and remove it from the @users array
        $pc1 = 0;
        $pc2 = $userLimit - $availPolicy.ExchangeLocation.count - 1
        $processUsers = $users[$pc1..$pc2]
        Add-UserToRetentionPolicy -PolicyName $availPolicy.Name -Users $processUsers
        write-host "$($processUsers.length) users were added to $($availPolicy.Name)."

        # removed processed users
        $pc1 = $userLimit - $availPolicy.ExchangeLocation.count
        $pc2 = $users.length - 1
        $users = $users[$pc1..$pc2]
    }
}


#batch add the rest of users to new retention policy
if ($users.length -gt 0) {
    $loopCount = [int]($users.length / $userLimit) + 1
    for ($index = 0; $index -lt $loopCount; $index++) {
        $ps = $userLimit * $index
        $pe = $userLimit * ($index + 1) - 1
        if ($pe -gt $users.length) {
            $pe = $users.length - 1
        }
    
        $processUsers = $users[$ps..$pe]
        
        $newPolicy = Create-RetentionPolicy -PolicyNamePrefix $pnPrefix -Users $processUsers
        write-host "$($processUsers.length) users were added to $($newPolicy.Name)."
    }
}

Disconnect-ExchangeOnline -Confirm:$false
write-host "Disconnected EXO online."


