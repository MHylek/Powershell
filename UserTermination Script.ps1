
$disableUser = 'ckapetan'
$forwardingAddress= 'mciesielska@epsteinglobal.com'
$cred = Get-Credential “Epstein\Administrator”

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://ChiExch01.corp.epstein-isi.com/PowerShell/ -Authentication Kerberos
$SFBsession = New-PSSession -ConnectionURI “https://chiskype01.corp.epstein-isi.com/OcsPowershell” -Credential $cred
$target = Get-ADOrganizationalUnit -LDAPFilter “(name=Disabled Accounts)”
#Create Group Logs
Get-ADPrincipalGroupMembership $disableUser | select name | Out-file \\chifile01\IT\FileDump\Terminations\$disableUser.txt
Get-ADPrincipalGroupMembership -Identity $disableUser| where {$_.Name -notlike "Domain Users"} |% {Remove-ADPrincipalGroupMembership -Identity $disableUser -MemberOf $_ -Confirm:$false}
Set-ADUser $disableUser -Description (Get-Date -Format MM-dd-yyyy) -OfficePhone " "
#Exchange
Import-PSSession $Session
Set-Mailbox $disableUser -ForwardingAddress  $ForwardingAddress -HiddenFromAddressListsEnabled $true
Disable-UMMailbox -Identity $disableUser -Confirm:$false
Send-MailMessage -From "Administrator@epsteinglobal.com" -To $forwardingAddress -bcc "MHylek@epsteinglobal.com" -Subject "Notice of mail forwarding for $disableuser " -Body "Hello, This email is a notification that you are being forwarded the emails of $disableUser . Normally we provide this service for 30 days, if you need this for longer than 30 days please submit a Helpdesk request " -SmtpServer Chimail01
Remove-PSSession $Session
#Skype For business
Import-PSSession $SFBsession
Disable-CsUser -Identity $disableUser
Remove-PSSession $SFBession
get-aduser $disableUser |move-ADObject  -targetpath $target
Disable-ADAccount $disableUser


Get-PSSession | Remove-PSSession
