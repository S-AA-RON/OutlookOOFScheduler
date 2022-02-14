if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
	Write-Host "ExchangeOnlineManagement exists, not installing`n"
} 
else {
	Write-Host "ExchangeOnlineManagement does not exist, installing`n"
	Import-Module ExchangeOnlineManagement -force
}
#this assumes you sign in with MS corp alias

$CurrentUser = ((Get-WMIObject -ClassName Win32_ComputerSystem).Username).Split('\')[1]
$CUAlias1 = $CurrentUser + "@Microsoft.com"
$CUAlias2 = $CurrentUser + "@MicrosoftSupport.com"

#connect to exchange
Connect-ExchangeOnline -UserPrincipalName $CUAlias1 #corp alias


$MailboxARC = Get-MailboxAutoReplyConfiguration -identity $CUAlias1 #copies corp auto reply configuration

####replace my corp alias here with your corp alias to change that part if present in your OOF message
$EXTMSG = $MailboxARC.ExternalMessage.replace("Aaron.Sanders@Microsoft.com",$CUAlias2)
$INTMSG = $MailboxARC.InternalMessage.replace("Aaron.Sanders@Microsoft.com",$CUAlias2)

Connect-ExchangeOnline -UserPrincipalName $CUAlias2 #support alias

Set-MailboxAutoReplyConfiguration –identity $CUAlias2 `
-ExternalMessage `
$EXTMSG -InternalMessage `
$INTMSG -StartTime `
$MailboxARC.StartTime -EndTime `
$MailboxARC.EndTime `
-AutoReplyState `
$MailboxARC.AutoReplyState 


#disconect PS exo connections x2
Disconnect-ExchangeOnline -Confirm:$false