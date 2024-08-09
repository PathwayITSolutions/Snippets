$LDAPDomain = "LDAP://DC=fabrikam,DC=com"
$LDAPAdminSDHolder = "LDAP://CN=AdminSDHolder,CN=System,DC=fabrikam,DC=com"
$LDAPDNS = "LDAP://CN=MicrosoftDNS,CN=System,DC=fabrikam,DC=com"

([adsisearcher]'(&(objectClass=user)(adminCount=1))').FindAll() >> C:\PathwayIT\MembersOfProtectedGroups.txt

([adsisearcher]'(&(objectCategory=user)(userAccountControl:1.2.840.113556.1.4.803:=524288))').FindAll() >> C:\PathwayIT\UnconstrainedDelegation.txt

$ADSI=[ADSI]"$($LDAPDomain)"
$ADSI.psbase.get_ObjectSecurity().getAccessRules($true, $true,[system.security.principal.NtAccount]) >> C:\PathwayIT\EnumerateACLsOnDNCObject.txt

$ADSI=[ADSI]"$($LDAPAdminSDHolder)"
$ADSI.psbase.get_ObjectSecurity().getAccessRules($true, $true,[system.security.principal.NtAccount]) >> C:\PathwayIT\EnumerateACLsOnAdminSDHolder.txt

$ADSI=[ADSI]"$($LDAPDNS)"
$ADSI.psbase.get_ObjectSecurity().getAccessRules($true, $true,[system.security.principal.NtAccount]) >> C:\PathwayIT\EnumerateACLsOnMicrosoftDNSContainer.txt

([adsisearcher]'(memberOf=cn=DnsAdmins,CN=Users,dc=fabrikam,dc=com)').FindAll() >> C:\PathwayIT\MembersOfDNSAdmins.txt

([adsisearcher]'(&(objectCategory=computer)(primaryGroupID=516))').FindAll() >> C:\PathwayIT\ListDomainControllers.txt

([adsisearcher]'(&(objectCategory=computer)(!(primaryGroupID=516)(userAccountControl:1.2.840.113556.1.4.803:=524288)))').FindAll() >> C:\PathwayIT\ServersWithUnconstrainedDelegation.txt
