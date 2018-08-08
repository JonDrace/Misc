###Define variables###
cls
$username = Read-Host -prompt 'Username'

###Open connection with exchange###

$SessionOpt = New-PSSessionOption -SkipCACheck:$true -SkipCNCheck:$true -SkipRevocationCheck:$true
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://DEUSEFRAN1502/powershell/ -AllowRedirection -SessionOption $SessionOpt
Import-PSSession $Session -AllowClobber

$mailbox_reply_config = Get-MailboxAutoReplyConfiguration –identity $username
cls

###Current state###
if (Get-Mailbox –identity $username -EA silentlycontinue){
	if ($mailbox_reply_config -ne $NULL) { #Check is any config/mailbox has been received
    	if ($mailbox_reply_config.Autoreplystate.ToString() -eq "Enabled" ) {
        	Write-Host("InternalMessage:")
        	$mailbox_reply_config.InternalMessage
        	Write-Host("ExternalMessage:")
        	$mailbox_reply_config.ExternalMessage
    	}
    	else {
        	Write-Host("Auto reply is currently: " + $mailbox_reply_config.Autoreplystate)
	        
    	}
	
	###State switches###
	
    	$reply_status_switch = Read-Host -prompt "E - enable`nD - disable`nDefault: Skip`n"
	
    	###Enable AutoReply###
    	if ($reply_status_switch -eq "E" -or $reply_status_switch -eq "e") {
        	$reply_status_internal = ( (Read-Host -prompt "Provide message for internal users").Replace("`n","<BR>") )
        	$reply_status_external = ( (Read-Host -prompt "Provide message for external users").Replace("`n","<BR>") )
        	Set-MailboxAutoReplyConfiguration $username –AutoReplyState Enabled –ExternalMessage $reply_status_external –InternalMessage $reply_status_internal
        	$mailbox_reply_config = Get-MailboxAutoReplyConfiguration –identity $username
    	}
    	###Disable AutoReply###
    	if ($reply_status_switch -eq "D" -or $reply_status_switch -eq "d") {
         	Set-MailboxAutoReplyConfiguration $username –AutoReplyState Disabled –ExternalMessage $null –InternalMessage $null
         	$mailbox_reply_config = Get-MailboxAutoReplyConfiguration –identity $username
    	}
	
    	###Last AutoReply state check###
	
    	Write-Host ("Closing connection`nCurrent reply config status: " + $mailbox_reply_config.Autoreplystate)
	}
	else {
    	Write-Host("No Auto reply configuration has been found for provided username") 
	}
}
else{
    Write-Host("User does not exist") 
}

###Close connection with exchange###

Remove-PSSession $Session