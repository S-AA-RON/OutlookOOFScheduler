$Global:UserAlias=
$Global:CurrentUser=
$Global:UserAliasSuffix="@MicrosoftSupport.com"
$Global:MailboxARC=
$Global:EndOfShift=
$Global:StartOfShift=
get-Alias
$AliasPath = $Global:UserAlias.replace("@","_")
$Global:MessageFilePath= "C:\Users\${Global:CurrentUser}\OneDrive - Microsoft\Desktop\oof message script\$AliasPath\"
ConnectAlias2EXO
#DisconnectEXO
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
	CurrentUserNamefromWindows
	
	$PromptText = "$Global:CurrentUser please enter the Alias Suffix of the Account to change. Ex. $Global:UserAliasSuffix"

	$Global:UserAliasSuffix = Read-Host -Prompt $prompttext
    if($Global:UserAliasSuffix -eq ""){ #if user doesn't input anything use default
		$Global:UserAliasSuffix="@MicrosoftSupport.com"
	}
    
    $Global:UserAlias = "$Global:CurrentUser$Global:UserAliasSuffix"
    Write-Host "UserAlias is $Global:UserAlias"
    Write-Host "UserAliasSuffix is $Global:UserAliasSuffix"
}

function ConnectAlias2EXO {
	InstallEXOM #is EXO module installed
	
	Write-Host "Connecting to your Outlook Account $UserAlias`n" 
	Connect-ExchangeOnline -UserPrincipalName $UserAlias
	Write-Host "Done Connecting"
}

function get-ARC {
    $TempPath = $Global:MessageFilePath + "AutoReplyConfig.json"
	if(Check-File($TempPath)) {
        Write-Host "AutoConfig has pre-existing file $TempPath"
        get-ARCFile
	}
    else {
         Write-Host "AutoConfig is being written to JSON file $TempPath"
	    $Global:MailboxARC = Get-MailboxAutoReplyConfiguration -identity $UserAlias
        $Global:MailboxARC | ConvertTo-Json -depth 100 | Set-Content $TempPath
    }
		
    $temp = $Global:MailboxARC.AutoReplyState
	Write-Host "Current Auto Reply State is : $temp"
}

function get-ARCFile {
   
    $TempPath = $Global:MessageFilePath + "AutoReplyConfig.json"
    $Global:MailboxARC = Get-Content $TempPath -Raw | ConvertFrom-Json 
}

function writeMessage {
    get-ARCFile
    #write external to html remove the'?' at the start, fix this bug later lol
    #overwrites from json
    #add file check
    $TempPath = $Global:MessageFilePath + "External.html"
    $Global:MailboxARC.ExternalMessage.substring(1) | Out-File -FilePath $TempPath
    #same but internal
    #overwrites from json
    #add file check
    $TempPath = $Global:MessageFilePath + "Internal.html"
    $Global:MailboxARC.InternalMessage.substring(1) | Out-File -FilePath $TempPath
}

####check file does exist
function Check-File($FilePath) {
    return (Get-Item -Path $FilePath -ErrorAction Ignore)
}

#set autoreply to scheduled
#this requires start and end times
function Set-ARCSTATEScheduled {
	if($Global:MailboxARC -eq $null){
		$Global:MailboxARC = get-arc
	}
	if($Global:UserAlias -eq $null){
		$Global:UserAlias = get-Alias
	}
	if($Global:StartOfShift -eq $null){
		$Global:StartOfShift = GetShiftTime("start")
	}
	if($Global:EndOfShift -eq $null){
		$Global:EndOfShift = GetShiftTime("end")
	}
	#is Reply state disabled or enabled by the user manually instead of scheduled
	if($Global:MailboxARC.AutoReplyState -eq "Disabled" -or $Global:MailboxARC.AutoReplyState -eq "Enabled"){
		$CurrentTime = Get-Date
		IsOfficeHours
	}
    #remove this else? this is the point of this function
    #add read from ARC json
	else {
		$Global:StartOfShift = $Global:StartOfShift.TimeofDay.AddDays(1)
		Set-MailboxAutoReplyConfiguration –identity $UserAlias `
		-ExternalMessage $Global:MailboxARC.ExternalMessage `
		-InternalMessage $Global:MailboxARC.InternalMessage `
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
    #read from stored file first then get from active and compare?
	if($Global:MailboxARC.ExternalMessage -eq $null -or $Global:MailboxARC.InternalMessage -eq $null ){
		get-ARC
	}
	if($Global:MailboxARC.ExternalMessage -and $Global:MailboxARC.InternalMessage) {
        #If messages are the same only write one message file to disk
		if($Global:MailboxARC.ExternalMessage -eq $Global:MailboxARC.InternalMessage){
            $TempPath = $Global:MessageFilePath + "External.html"
            Write-Host "The internal and external messages are the same. `nOne OOF Message to Rule them All `n Writing Message to HTML $TempPath"
            $Global:MailboxARC.ExternalMessage.substring(1) | Out-File -FilePath $TempPath
		}
		else{
			Write-Host "Differenet External and Internal Messages"

            $TempPath = $Global:MessageFilePath + "External.html"
            $Global:MailboxARC.ExternalMessage.substring(1) | Out-File -FilePath $TempPath
 
            $TempPath = $Global:MessageFilePath + "Internal.html"
            $Global:MailboxARC.InternalMessage.substring(1) | Out-File -FilePath $TempPath
		}
	}
}

function Set-Message {
	$MessageFilePath = "C:\Users\${Global:CurrentUser}\OneDrive - Microsoft\Desktop\oof message script\$({Global:UserAlias}.replace("@","_"))\External.html"

	#this IF assumes if there is only 1 message file the messages are the same
	#save as HTML for better format editing by end user
	#check of separate files AND/OR save them in 1 file and be smart about reading it
	
	if(Check-File($TempPath)){ 
		Write-Host "Setting the same OOF Message for both Internal and External"
		$Message = [System.IO.File]::ReadAllText($MessageFilePath)
		Write-Output $Message
		Set-MailboxAutoReplyConfiguration –identity $Global:UserAlias –ExternalMessage $Message -InternalMessage $Message
	}
	else{  
		Write-Host "Different External and Internal Messages"
				
		$Message = [System.IO.File]::ReadAllText($MessageFilePath)
		Write-Output ("Setting External Message`n`n" + $Message)
		Set-MailboxAutoReplyConfiguration –identity $UserAlias –ExternalMessage $Message
		
		$MessageFilePath = $MessageFilePath -replace 'Ex','In' #only change in file names is Ex to In
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
