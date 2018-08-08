function Open_Connection_to_Exchange(){
	param(
		[Parameter(Mandatory = $false)]
		[bool]$specifyCredentials
	)
	if($specifyCredentials){
		$LiveCred = Get-Credential -Credential "@grouphc.net"
		$SessionOpt = New-PSSessionOption -SkipCACheck:$true -SkipCNCheck:$true -SkipRevocationCheck:$true
		$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://DEUSEFRAN1502/powershell/ -Credential $LiveCred -AllowRedirection -SessionOption $SessionOpt 
		Import-PSSession $Session
	}
	else{
		$SessionOpt = New-PSSessionOption -SkipCACheck:$true -SkipCNCheck:$true -SkipRevocationCheck:$true
		$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://DEUSEFRAN1505/powershell/ -AllowRedirection -SessionOption $SessionOpt 
		Import-PSSession $Session -AllowClobber
	} 
}

function Load_List(){
	param(
		[Parameter(Mandatory=$False)]
		$promptMessage = "Add item"
	)
	Write-Output ("Press Enter on empty value to end cycle")
	while ($loaded_list_item = Read-Host -prompt $promptMessage) {[array]$loaded_list += $loaded_list_item}
	return $loaded_list
}

function Get_Save_Location(){
	$Destination_Path = New-Object -Typename System.Windows.Forms.SaveFileDialog
	$Destination_Path.Filter = "All files (*.*) | *.*"
	$Destination_Path.ShowDialog()
	return $Destination_Path.FileName
}

function Add_Mailbox_Permission(){
	param(
		[Parameter(Mandatory = $false)][bool]$fullmailboxaccess=$false,
		[Parameter(Mandatory = $false)][bool]$automap=$false,
		[Parameter(Mandatory = $false)][bool]$sendonbehalf=$false,
		[Parameter(Mandatory = $false)][bool]$sendas=$false,
		[Parameter(Mandatory = $true)][string]$requester,
		[Parameter(Mandatory = $true)][string]$targetMailbox
	)

	#Full Mailbox Access
	if ($fullmailboxaccess) {
		if ($automap) {
			Add-MailboxPermission -Identity $targetMailbox -User $requester -AccessRights FullAccess -AutoMapping $true
		}
		else {
			Add-MailboxPermission -Identity $targetMailbox -User $requester -AccessRights FullAccess -AutoMapping $false
		}
	}
	
	#Send On Behalf
	if ($sendonbehalf) {
		Set-Mailbox $targetMailbox -GrantSendOnBehalfTo @{Add=$requester}
	}

	#Send As
	if ($sendas) {
		Add-ADPermission -Identity $targetMailbox -User $requester -Extendedrights "Send As"
	}	
}