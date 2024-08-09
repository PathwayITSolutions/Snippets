$Credential = Get-Credential
Connect-MsolService -Credential $Credential
Connect-SPOService -Credential $Credential -Url https://contoso-admin.sharepoint.com

$list = New-Object System.Collections.ArrayList
#Counters
$i = 0


#Get licensed users
$users = Get-MsolUser -All | Where-Object { $_.islicensed -eq $true }
#total licensed users
$count = $users.count

foreach ($u in $users) {
    $i++
    Write-Host "$i/$count"

    $upn = $u.userprincipalname
    $list.Add($upn) | Out-Null

    if ($i -eq 199) {
        #We reached the limit
        Request-SPOPersonalSite -UserEmails $list -NoWait
        Start-Sleep -Milliseconds 655
        $list.Clear()
        $i = 0
    }
}

if ($i -gt 0) {
    Request-SPOPersonalSite -UserEmails $list -NoWait
}
