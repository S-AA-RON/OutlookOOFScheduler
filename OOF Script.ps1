function InstallEXOM {
	if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
		Write-Host "ExchangeOnlineManagement exists, not installing`n"
	} 
	else {
		Write-Host "ExchangeOnlineManagement does not exist, installing`n"
		Import-Module ExchangeOnlineManagement -force
	}
}

function CurrentUserNamefromWindows {
	$CurrentUser = ((Get-WMIObject -ClassName Win32_ComputerSystem).Username).Split('\')[1]
	return $CurrentUser
}

function get-Alias {
	if($CurrentUser -eq $undefinedVariable){
		$CurrentUser = CurrentUserNamefromWindows
	}
	$PromptText = "Enter the Alias Suffix you want to change. Ex. @MicrosofSupport"
	$UserAliasSuffix = Read-Host -Prompt $prompttext
	if($UserAliasSuffix -eq $undefinedVariable){
		$UserAliasSuffix = "@MicrosoftSupport.com"
	}
	$UserAlias = $CurrentUser + $UserAliasSuffix
	return $UserAlias
}

function ConnectAlias2EXO {
	Write-Host "Connecting to your Outlook Account`n"
	if($UserAlias -eq $undefinedVariable){
	$UserAlias = get-Alias(CurrentUserNamefromWindows)   
	}
	Connect-ExchangeOnline -UserPrincipalName $UserAlias
	Write-Host "Done Connecting"
}

function get-arc {
	if($UserAlias -eq $undefinedVariable){
		$UserAlias = get-Alias
	}
	$MailboxARC = Get-MailboxAutoReplyConfiguration -identity $UserAlias
	return $MailboxARC
}

function GET-ARCSTATE {
	$MailboxARC = get-arc
	Write-Host "Current Auto Reply State is :" + $MailboxARC.AutoReplyState
	return $MailboxARC.AutoReplyState
}

function Set-ARCSTATEScheduled {
	if($MailboxARC -eq $undefinedVariable){
		$MailboxARC = get-arc
	}
	if($UserAlias -eq $undefinedVariable){
		$UserAlias = get-Alias
	}
	if($StartofShift -eq $undefinedVariable){
		$StartofShift = GetShiftTime("start")
	}
	if($EndofShift -eq $undefinedVariable){
		$EndofShift = GetShiftTime("end")
	}
	if($MailboxARC.AutoReplyState -eq "Disabled"-or $MailboxARC.AutoReplyState -eq "Enabled"){
		$CurrentTime = Get-Date
		if($CurrentTime.TimeOfDay -lt $StartofShift.TimeOfDay){ 
			Write-Host "Currently Before Shift`n"
		}
		elseif($CurrentTime.TimeOfDay -gt $EndofShift.TimeOfDay){
			Write-Host "Currently After Shift`n"
		}
		elseif($EndofShift.TimeOfDay -le $CurrentTime.TimeOfDay -And $CurrentTime.TimeOfDay -ge $StartofShift.TimeOfDay){
			Write-Host "Currently During Shift`n"
		}
		else {
			Write-Host "Twilight Zone"
		}
	}
	$StartofShift = $StartofShift.TimeofDay.AddDays(1)
	Set-MailboxAutoReplyConfiguration –identity $UserAlias `
	-ExternalMessage $MailboxARC.ExternalMessage `
	-InternalMessage $MailboxARC.InternalMessage `
	-StartTime $EndofShift.TimeofDay `
	-EndTime $StartofShift.TimeofDay `
	-AutoReplyState "Scheduled"
}

function Get-Message {
	if($CurrentUser -eq $undefinedVariable){
		$CurrentUser = CurrentUserNamefromWindows
	}
	if($MailboxARC.ExternalMessage -eq $undefinedVariable){
		$MailboxARC = get-arc
	}

	$MessageFilePath = "C:\Users\" + $CurrentUser + "\OneDrive - Microsoft\Desktop\oof message script\OOFMessage"

	if($MailboxARC.ExternalMessage -and $MailboxARC.InternalMessage) {
		if($MailboxARC.ExternalMessage -eq $MailboxARC.InternalMessage){
			Write-Host "The internal and external messages are the same. `nOne OOF Message to Rule them All `n"
			$MailboxARC.ExternalMessage | Out-File ($MessageFilePath + ".txt")
			Write-Output $MailboxARC.ExternalMessage
		}
		else{
			Write-Host "Differenet External and Internal Messages"

			$MailboxARC.ExternalMessage | Out-File ($MessageFilePath + "_External.txt")
			Write-Output $MailboxARC.ExternalMessage

			$MailboxARC.InternalMessage | Out-File ($MessageFilePath + "_Internal.txt")
			Write-Output $MailboxARC.InternalMessage
		}
	}
}

function Set-Message {
	if($CurrentUser -eq $undefinedVariable){
		$CurrentUser = CurrentUserNamefromWindows
	}
	$MessageFilePath = "C:\Users\" + $CurrentUser + "\OneDrive - Microsoft\Desktop\oof message script\OOFMessage"
	if($UserAlias -eq $undefinedVariable){
		$UserAlias = get-Alias
	}
	#this IF assumes if there is only 1 message file the messages are the same
	#save as HTML for better format editing by end user
	#check of separate files AND/OR save them in 1 file and be smart about reading it
	
	if(($MessageFilePath + ".txt")){ 
		Write-Host "Setting the same OOF Message for both Internal and External"
		$Message = [System.IO.File]::ReadAllText($MessageFilePath+".txt")
		Write-Output $Message
		Set-MailboxAutoReplyConfiguration –identity $UserAlias –ExternalMessage $Message -InternalMessage $Message
	}
	else{  
		Write-Host "Different External and Internal Messages"
		$MessageFilePath = $MessageFilePath + "_External.txt" #external file name
		
		$Message = [System.IO.File]::ReadAllText($MessageFilePath)
		Write-Output ("Setting External Message`n`n" + $Message)
		Set-MailboxAutoReplyConfiguration –identity $UserAlias –ExternalMessage $Message
		
		$MessageFilePath = $MessageFilePath - "Ex" + "In" #only change in file names is Ex to In
		$Message = [System.IO.File]::ReadAllText($MessageFilePath)
		Write-Output ("Setting Internal Message`n`n" + $Message)
		Set-MailboxAutoReplyConfiguration –identity $UserAlias -InternalMessage $Message
	}
}

function GetShiftTime($StartEnd) { 
	$PromptText = "Enter when you " + $StartEnd + " your work day. Ex 9:00am"
	$ShiftTime = Read-Host -Prompt $PromptText
	$ShiftTimeOut = [datetime] $ShiftTime
	return $ShfitTimeOut
}

function DisconnectEXO {
	Disconnect-ExchangeOnline -Confirm:$false
}
