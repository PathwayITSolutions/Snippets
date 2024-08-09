#Module check for Active Directory and Exchange

if(!(Get-Module ActiveDirectory)) {
    Write-Host "Active Directory Module required"
    exit
}
if(!(Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn)) {
    Write-Host "Exchange Snap In required"
    exit
}

#Credential check for Domain Admin and Exchange Organization Management role

$CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$WindowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal($CurrentUser)
if(!($WindowsPrincipal.IsInRole("Domain Admins"))) {
       Write-Host "This script must be run with Domain Administrator credentials" -ForegroundColor Red
       Write-Host "Please open PowerShell with Domain Administrator credentials and run script again"
       exit
    }

    if(!($WindowsPrincipal.IsInRole("Exchange Organization Administrators"))) {
        Write-Host "This script must be run with Exchange Organization Administrator credentials" -ForegroundColor Red
        Write-Host "Please open PowerShell with Exchange Organization Administrator credentials and run script again"
        exit
     }

Set-AdServerSettings -ViewEntireForest $True

$SourceArray = New-Object System.Collections.ArrayList
$DestinationArray = New-Object System.Collections.ArrayList

$DistrGroups = Get-DistributionGroup -resultsize unlimited | Where-Object { ($_.identity -like "*") <#-and ($_.identity -notlike "* *") -and ($_.identity -notlike "* *")#> }
Foreach ($DistrGroup in $DistrGroups) {
    $SourceArray = "" | Select-Object "Alias", "DisplayName", "EmailAddress", "PrimarySmtpAddress", "Guid", "SamAccountName", "EmailAddresses"
    $SourceArray.Alias = $DistrGroup.Alias
    $SourceArray.DisplayName = $DistrGroup.DisplayName
    $SourceArray.EmailAddress = $DistrGroup.WindowsEmailAddress
    $SourceArray.PrimarySmtpAddress = $DistrGroup.PrimarySmtpAddress
    $SourceArray.Guid = $DistrGroup.Guid
    $SourceArray.SamAccountName = $DistrGroup.SamAccountName
    $SourceArray.EmailAddresses = $DistrGroup.EmailAddresses

    Try {
        If (Get-ADGroup $SourceArray.SamAccountName) {
            Write-Output "$($SourceArray.DisplayName) Destination AD Group discovered" >> c:\scripts\log.txt
            Write-Output "$($SourceArray.Alias)" >> c:\scripts\log.txt
            Write-Output "$($SourceArray.DisplayName)" >> c:\scripts\log.txt
            Write-Output "$($SourceArray.EmailAddress)" >> c:\scripts\log.txt
            Write-Output "$($SourceArray.PrimarySmtpAddress)" >> c:\scripts\log.txt
            Write-Output "$($SourceArray.Guid)" >> c:\scripts\log.txt
            Write-Output "$($SourceArray.SamAccountName)" >> c:\scripts\log.txt
            foreach ( $EmailAddress in $($SourceArray.EmailAddresses) ) {
                Write-Output "$($EmailAddress.ProxyAddressString)" >> c:\scripts\log.txt
            }
            Disable-DistributionGroup $SourceArray.SamAccountName
            Write-Output "$($SourceArray.DisplayName) Source Distribution Group disabled" >> c:\scripts\log.txt
            $DNStoreObj = Get-Adgroup $SourceArray.SamAccountName
            $DestinationArray = "" | Select-Object "DistinguishedName", "ObjectGUID"
            $DestinationArray.DistinguishedName = $DNStoreObj.DistinguishedName
            $DestinationArray.ObjectGUID = $DNStoreObj.ObjectGUID
            Enable-DistributionGroup -Identity $DestinationArray.DistinguishedName -PrimarySmtpAddress $SourceArray.PrimarySmtpAddress
            Write-Output "-->" >> c:\scripts\log.txt
            Write-Output "$($DestinationArray.DistinguishedName)" >> c:\scripts\log.txt
            Write-Output "$($DestinationArray.ObjectGUID)" >> c:\scripts\log.txt
            Write-Output "" >> c:\scripts\log.txt
        }
    }
    Catch {
        Write-Output "$($SourceArray.DisplayName) Destination AD Group doesn't exist" >> c:\scripts\log.txt
        Write-Output "$($SourceArray.Alias)" >> c:\scripts\log.txt
        Write-Output "$($SourceArray.DisplayName)" >> c:\scripts\log.txt
        Write-Output "$($SourceArray.EmailAddress)" >> c:\scripts\log.txt
        Write-Output "$($SourceArray.PrimarySmtpAddress)" >> c:\scripts\log.txt
        Write-Output "$($SourceArray.Guid)" >> c:\scripts\log.txt
        Write-Output "$($SourceArray.SamAccountName)" >> c:\scripts\log.txt
        foreach ($EmailAddress in $SourceArray.EmailAddresses) {
            Write-Output "$($EmailAddress.ProxyAddressString)" >> c:\scripts\log.txt
        }
        Write-Output "-->" >> c:\scripts\log.txt
        Write-Output "$($DestinationArray.DistinguishedName)" >> c:\scripts\log.txt
        Write-Output "$($DestinationArray.ObjectGUID)" >> c:\scripts\log.txt
        Write-Output "" >> C:\scripts\log.txt
        <# New-ADGroup -Name -SamAccountName "$($SourceArray.SamAccountName)DIST" -GroupCategory Distribution -GroupScope Global -DisplayName "$($SourceArray.DisplayName)DIST" -Path "ou=Temp,ou=Distribution,ou=Groups,ou=US,ou=Managed,dc=craneae,dc=com"
        Get-ADGroupMember -Identity $($SourceArray.Guid)
        foreach {Add-ADGroupMember -Identity XXXX -Members $($_.DistinguishedName)} #>
    }
}
