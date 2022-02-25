$Global:UserAlias=
$Global:CurrentUser=
$Global:UserAliasSuffix="@MicrosoftSupport.com"
$Global:MailboxARC=

##comment this out if you want it to always ask for time input when run
$Global:EndOfShift= [datetime] "6:00pm"
$Global:StartOfShift= [datetime] "9:00am"
## if commenting the above two lines out, remove the # from the next two lines
#$Global:EndOfShift=
#$Global:StartOfShift=

get-Alias
ConnectAlias2EXO
get-ARC
writeARC2File
get-Message
writemessage
#CheckStartEnd
set-ARCSTATEScheduled
DisconnectEXO

function CreateOOFPath {
    If(!(test-path $Global:MessageFilePath))
    {
          New-Item -ItemType Directory -Force -Path $Global:MessageFilePath
    }
}

$AliasPath = $Global:UserAlias.replace("@","_")
$AliasPath = $AliasPath.replace(".","_")
$Global:MessageFilePath= "C:\Users\${Global:CurrentUser}\OneDrive - Microsoft\Desktop\oof message script\$AliasPath\"
CreateOOFPath

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
    $Global:MailboxARC = Get-MailboxAutoReplyConfiguration -identity $UserAlias
	if(Check-File($TempPath)) {
        Write-Host "AutoConfig has pre-existing file $TempPath"
        $JSONMailboxARC = Get-Content $TempPath -Raw | ConvertFrom-Json 
        #compare file and current configuration
        #ask to use one or the other
        if($Global:MailboxARC -eq $JSONMailboxARC){
            Write-Host "JSON is same as current config"
        }
        else {
            Write-Host "JSON Differs"
        }
            
	}
    else {
        Write-Host "AutoConfig is being written to JSON file $TempPath"
        $Global:MailboxARC | ConvertTo-Json -depth 100 | Set-Content $TempPath
    }
		
    $temp = $Global:MailboxARC.AutoReplyState
	Write-Host "Current Auto Reply State is : $temp"
}

function writeARC2File {
    $TempPath = $Global:MessageFilePath + "AutoReplyConfig.json"
    Write-Host "Writing Mailbox Auto Reply to JSON file $TempPath"
    $Global:MailboxARC = Get-Content $TempPath -Raw | ConvertFrom-Json 
}

function writeMessage {
    writeARC2File
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

####get start and end shift times
### returns [datetime]
function GetShiftTime($StartEnd) { 
	$PromptText = "Enter when you " + $StartEnd + " your work day. Ex 9:00am"
	$ShiftTime = Read-Host -Prompt $PromptText
	$ShiftTimeOut = [datetime] $ShiftTime
	return $ShfitTimeOut
}

###check stored start and end times 
function CheckStartEnd {
    <#

    #######################JSON TIME DOES NOT WORK///
    #are there start and end times in current json file?
    $TempPath = $Global:MessageFilePath + "AutoReplyConfig.json"
    #does the json exist
	if(Check-File($TempPath)) {
        Write-Host "AutoConfig JSON exist at $TempPath"
        $JSONMailboxARC = Get-Content $TempPath -Raw | ConvertFrom-Json 
    }
    if($JSONMailboxARC.StartTime -ne $null -and $JSONMailboxARC.EndTime -ne $null) {
        $tempstarttime = [datetime] ($JSONMailboxARC.StartTime.TimeOfDay)
        $tempendtime = [datetime] ($JSONMailboxARC.EndTime.TimeOfDay)
        $PromptText = "Does your shift end today at $tempstarttime and start tomorrow at $tempendtime? [Y]es/[N]o"
        $YesNo = Read-Host -Prompt $PromptText
        if($YesNo -ne "N" -or $YesNo -ne "n" -or $YesNo -eq "") {
            #add a day for tomorrow store in global values as user did not want to change them

            $temptime = [datetime] ($JSONMailboxARC.StartTime.TimeOfDay)
            $Global:MailboxARC.EndTime = $temptime.AddDays(1)
            $Global:MailboxARC.StartTime = [datetime]$JSONMailboxARC.EndTime.TimeOfDay
        }
        else {
            #user does not like json values
            #or one JSON time is Null
            #then ask the user for their input

            #is not null, then is it correct?
            if($JSONMailboxARC.EndTime -ne $null) {
                $PromptText = "Does your shift start at ${JSONMailboxARC.EndTime.TimeOfDay}? [Y]es/[N]o"
                $YesNo = Read-Host -Prompt $PromptText
                if($YesNo -eq "N" -or $YesNo -eq "n" -or $YesNo -ne "") {
                    $Global:StartOfShift = GetShiftTime("start")
                    #add a day for tomorrow                    
                }
                $Global:MailboxARC.EndTime = [datetime]$Global:StartOfShift.TimeofDay.AddDays(1)
            }
            #if null ask user right away
            else {
                $Global:StartOfShift = GetShiftTime("start")                
            }
            #store start time in global
            $Global:MailboxARC.EndTime = [datetime]$Global:StartOfShift.TimeofDay.AddDays(1)

            #if not null, is it correct
            if($JSONMailboxARC.EndTime -ne $null) {
                $PromptText = "Does your shift end at ${JSONMailboxARC.StartTime.TimeOfDay}? [Y]es/[N]o"
                $YesNo = Read-Host -Prompt $PromptText
	            if($YesNo -eq "N" -or $YesNo -eq "n" -or $YesNo -ne "") {
		            $Global:EndOfShift = GetShiftTime("end")
                }
            }
            #if null ask user right away
            else {
                $Global:EndOfShift = GetShiftTime("end")                
            }
            #store start time in global
            $Global:MailboxARC.StartTime = [datetime]$Global:EndOfShift.TimeofDay

        }
        #store global values in json, regardless of change
        $Global:MailboxARC | ConvertTo-Json -depth 100 | Set-Content $TempPath
    }
    #>
    #are there start and end times in script file? 
    if($Global:StartOfShift -ne $null -and $Global:EndOfShift -ne $null) {
        $temptimemath = Get-Date($Global:StartOfShift).AddDays(1)
        $PromptText = "PS VARs Does your shift end today at $Global:EndOfShift and start tomorrow at $temptimemath ? [Y]es/[N]o "
        $YesNo = Read-Host -Prompt $PromptText
        if($YesNo -eq "N" -or $YesNo -eq "n") {
            #check to see if start is already correctly configured
            $PromptText = "Does your shift start at $Global:StartOfShift [Y]es/[N]o"
            $YesNo = Read-Host -Prompt $PromptText
            if($YesNo -eq "N" -or $YesNo -eq "n" -and $YesNo -ne "") {
                $Global:StartOfShift = GetShiftTime("start")
                #add a day for tomorrow, store in global arc config
            }	    
           
            $PromptText = "Does your shift end at $Global:EndOfShift [Y]es/[N]o"
            $YesNo = Read-Host -Prompt $PromptText
            if($YesNo -eq "N" -or $YesNo -eq "n" -and $YesNo -ne "") {
                $Global:EndOfShift = GetShiftTime("end")
     
            }	    
        #store in global arc config
        $Global:StartOfShift = $temptimemath
        $Global:MailboxARC.EndTime = $Global:StartOfShift
        $Global:MailboxARC.StartTime = $Global:EndOfShift
        }
    }
    <#
    #by this point there should be values for both but why not check
    if($Global:MailboxARC.StartTime -eq $null -or $Global:MailboxARC.EndTime -eq $null) {
        #if there are times in Global ARC are they correct
        $PromptText = "Global Does your shift end today at $Global:MailboxARC.StartTime and start tomorrow at $Global:MailboxARC.EndTime? [Y]es/[N]o"
        $YesNo = Read-Host -Prompt $PromptText
        if($YesNo -ne "N" -or $YesNo -ne "n" -or $YesNo -eq "") {
            #add a day for tomorrow store in global values as user did not want to change them
            $Global:MailboxARC.EndTime = [datetime]$Global:MailboxARC.StartTime.TimeOfDay.AddDays(1)
            $Global:MailboxARC.StartTime = [datetime]$Global:MailboxARC.EndTime.TimeOfDay
        }
        #if values are not correct then ask the user for their input
        if($Global:MailBoxARC.EndTime -ne $null) {
            $PromptText = "Does your shift start at ${Global:MailBoxARC.EndTime}? [Y]es/[N]o"
            $YesNo = Read-Host -Prompt $PromptText
            #if not correct and not enter (blank)
            if($YesNo -eq "N" -or $YesNo -eq "n" -and $YesNo -ne "") {
                $Global:StartOfShift = GetShiftTime("start")
            }
        }
        #is null ask for input
        else {
            $Global:StartOfShift = GetShiftTime("start")
        }
        #add a day for tomorrow, store in global value
        $Global:MailboxARC.EndTime = [datetime]$Global:StartOfShift.TimeofDay.AddDays(1)

        if($Global:MailBoxARC.StartTime -ne $null) {       
            $PromptText = "Does your shift end at ${Global:MailBoxARC.StartTime}? [Y]es/[N]o"
            $YesNo = Read-Host -Prompt $PromptText
	        if($YesNo -eq "N" -or $YesNo -eq "n" -and $YesNo -ne "") {
		        $Global:EndOfShift = GetShiftTime("end")                
            }
        }
        #is null get input
        else {
            $Global:EndOfShift = GetShiftTime("end")
        }
        #store in global value
        $Global:MailboxARC.StartTime = [datetime]$Global:EndOfShift.TimeofDay
	} 
    #If either of the $Global:StartOfShift or $Global:EndOfShift are not configured, the above should ask for them#>
}

#set autoreply to scheduled
#this requires start and end times
function Set-ARCSTATEScheduled {
	#is Reply state disabled or enabled by the user manually instead of scheduled
    #remove this else? what is the point of this function
	IsOfficeHours
	
    #add read from ARC json

    #do you work tomorrow?

    #days of the week you work
    #store shedule in json file 
    #working days of the week [0,1,1,1,1,1,0]

    #store start and end times per day? per shift?
    
    if($Global:MailboxARC.AutoReplyState -eq "Scheduled") {

        #enable schedule yes/no
        $PromptText = "Would you like to enable a scheduled OOF message? [Y]es/[N]o"
	    $YesNo = Read-Host -Prompt $PromptText
	    if($YesNo -eq "Y" -or $YesNo -eq "y" -or $YesNo -eq "") {
            CheckStartEnd
            Set-MailboxAutoReplyConfiguration –identity $UserAlias `
		    -ExternalMessage $Global:MailboxARC.ExternalMessage `
		    -InternalMessage $Global:MailboxARC.InternalMessage `
		    -StartTime $Global:MailboxARC.StartTime.TimeofDay `
		    -EndTime $Global:MailboxARC.EndTime.TimeofDay `
		    -AutoReplyState "Scheduled"
        }
        else {
            Write-Host "Why run the command if you do not want to use it?"
        }
    }
}

function IsOfficeHours {
    $tempstate = $Global:MailboxARC.AutoReplyState
    $CurrentTime = Get-Date
    #$CurrentTime.TimeOfDay
    #$Global:MailboxARC.StartTime.TimeOfDay
    #$Global:MailboxARC.EndTime.TimeOfDay
	#check if it is during shift return bool based on start and end time
	if($CurrentTime.TimeOfDay -lt $Global:MailboxARC.EndTime.TimeOfDay){ 
		Write-Host "Before Shift and auto reply for $Global:UserAlias is $tempstate`n"
	}
	elseif($CurrentTime.TimeOfDay -gt $Global:MailboxARC.StartTime.TimeOfDay){
		Write-Host "After Shift and auto reply for $Global:UserAlias is $tempstate`n"
	}
	elseif($Global:MailboxARC.EndTime.TimeOfDay -le $CurrentTime.TimeOfDay -And $CurrentTime.TimeOfDay -ge $Global:MailboxARC.StartTime.TimeOfDay){
		Write-Host "During Shift and auto reply for $Global:UserAlias is $tempstate`n"
		return $true
	}
	else {
		Write-Host "Twilight Zone"
	}
	return $false
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
            Write-Host "The internal and external messages are the same.`nOne OOF Message to Rule them All `nWriting Message to HTML $TempPath"
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

function DisconnectEXO {
	Disconnect-ExchangeOnline -Confirm:$false
}
