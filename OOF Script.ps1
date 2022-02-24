$Global:UserAlias=
$Global:CurrentUser=
$Global:UserAliasSuffix="@MicrosoftSupport.com"
$Global:MailboxARC=
$Global:MessageFilePath= "C:\Users\$Global:CurrentUser\OneDrive - Microsoft\Desktop\oof message script\AutoReplyConfig.json"

function InstallEXOM {
	if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
		#Write-Host "ExchangeOnlineManagement exists, not installing`n"
        #no output if it is installed, less chatty
        return
	} 
	else {
		Write-Host "ExchangeOnlineManagement does not exist, installing`n"
		Import-Module ExchangeOnlineManagement -force
	}
}

#get current username from local user foldername
function CurrentUserNamefromWindows {
	$Global:CurrentUser = ((Get-WMIObject -ClassName Win32_ComputerSystem).Username).Split('\')[1]
    Write-Host "CurrentUser is $Global:CurrentUser"
}

function get-Alias {
	if($Global:CurrentUser -eq $null) {
		CurrentUserNamefromWindows
	}
	$PromptText = "Enter the Alias Suffix of the Account to change. Ex. $UserAliasSuffix"

	$Global:UserAliasSuffix = Read-Host -Prompt $prompttext
    if($Global:UserAliasSuffix -eq ""){ #if user doesn't input anything use default
		$Global:UserAliasSuffix="@MicrosoftSupport.com"
	}
    Write-Host "UserAliasSuffix is $Global:UserAliasSuffix"
	
    $Global:UserAlias = "$Global:CurrentUser$Global:UserAliasSuffix"
    Write-Host "UserAlias is $Global:UserAlias"
}

function ConnectAlias2EXO {
	InstallEXOM #is EXO module installed
	if($Global:UserAlias -eq $null){
		get-Alias
	} 
	Write-Host "Connecting to your Outlook Account $UserAlias`n" 
	Connect-ExchangeOnline -UserPrincipalName $UserAlias
	Write-Host "Done Connecting"
}


#####
function get-ARC {
	if($Global:CurrentUser -eq $null) {
		CurrentUserNamefromWindows
	}
    if($Global:UserAlias -eq $null){
		get-Alias
	}

    $Global:MessageFilePath = "C:\Users\$Global:CurrentUser\OneDrive - Microsoft\Desktop\oof message script\AutoReplyConfig.json"

	if(Check-File($Global:MessageFilePath)) {
        #read file here from json
	}
    else {
        ConnectAlias2EXO
	    $Global:MailboxARC = Get-MailboxAutoReplyConfiguration -identity $UserAlias
    }
	
	$Global:MailboxARC | ConvertTo-Json -depth 100 | Set-Content $Global:MessageFilePath

	Write-Host "Current Auto Reply State is : 'n" + (get-ARCState)
}


####check file does exist
function Check-File($FilePath) {
    return (Get-Item -Path $FilePath -ErrorAction Ignore)
	
}

#set autoreply to scheduled
#this r=uires start and end times
function Set-ARCSTATEScheduled {
	if($MailboxARC = $null){
		$MailboxARC = get-arc
	}
	if($UserAlias = $null){
		$UserAlias = get-Alias
	}
	if($StartOfShift = $null){
		$StartOfShift = GetShiftTime("start")
	}
	if($EndOfShift = $null){
		$EndOfShift = GetShiftTime("end")
	}
	#is Reply state disabled or enabled by the user manually instead of scheduled
	if($MailboxARC.AutoReplyState -eq "Disabled" -or $MailboxARC.AutoReplyState -eq "Enabled"){
		$CurrentTime = Get-Date
		IsOfficeHours
	}
	else { #set to scheduled
		$StartOfShift = $StartOfShift.TimeofDay.AddDays(1)
		Set-MailboxAutoReplyConfiguration –identity $UserAlias `
		-ExternalMessage $MailboxARC.ExternalMessage `
		-InternalMessage $MailboxARC.InternalMessage `
		-StartTime $EndOfShift.TimeofDay `
		-EndTime $StartOfShift.TimeofDay `
		-AutoReplyState "Scheduled"
	}
}

function IsOfficeHours {
	#check if it is during shift return bool based on start and end time
	if($CurrentTime.TimeOfDay -lt $StartOfShift.TimeOfDay){ 
		Write-Host "Currently Before Shift`n"
	}
	elseif($CurrentTime.TimeOfDay -gt $EndOfShift.TimeOfDay){
		Write-Host "Currently After Shift`n"
	}
	elseif($EndOfShift.TimeOfDay -le $CurrentTime.TimeOfDay -And $CurrentTime.TimeOfDay -ge $StartOfShift.TimeOfDay){
		Write-Host "Currently During Shift`n"
		Return True
	}
	else {
		Write-Host "Twilight Zone"
	}
	Return False
}

function Get-Message {
	if($Global:CurrentUser = $null){
		CurrentUserNamefromWindows
	}
    #read from stored file first then get from active and compare?
	if($Global:MailboxARC.ExternalMessage = $null){
		get-arc
	}

	$Global:MessageFilePath = "C:\Users\${Global:CurrentUser}\OneDrive - Microsoft\Desktop\oof message script\OOFMessage"

	if($Global:MailboxARC.ExternalMessage -and $MailboxARC.InternalMessage) {
		if($Global:MailboxARC.ExternalMessage -= $MailboxARC.InternalMessage){
			Write-Host "The internal and external messages are the same. `nOne OOF Message to Rule them All `n"
			$Global:MailboxARC.ExternalMessage | Out-File ($Global:MessageFilePath + ".txt")
			Write-Output $Global:MailboxARC.ExternalMessage
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
	if($Global:CurrentUser = $null){
		CurrentUserNamefromWindows
	}
	$MessageFilePath = "C:\Users\$Global:CurrentUser\OneDrive - Microsoft\Desktop\oof message script\OOFMessage"
	if($UserAlias = $null){
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
