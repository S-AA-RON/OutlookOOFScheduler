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
	$prompttext = "Enter the Alias Suffix you want to change. Ex. @MicrosofSupport"
	$userAliasSuffix = Read-Host -Prompt $prompttext
	$userAliasSuffix = "@MicrosoftSupport.com"
 if($CurrentUser -eq $undefinedVariable){
        $CurrentUser = CurrentUserNamefromWindows
    }
	$userAlias = $CurrentUser + $userAliasSuffix
    return $useralias
}

function ConnectAlias2EXO {
    Write-Host "Connecting to your Outlook Account`n"
    if($useralias -eq $undefinedVariable){
	$useralias = get-Alias(CurrentUserNamefromWindows)   
    }
    Connect-ExchangeOnline -UserPrincipalName $useralias
    Write-Host "Done Connecting"
}

function get-arc {
    if($useralias -eq $undefinedVariable){
        $useralias = get-Alias
    }
    $MailboxARC = Get-MailboxAutoReplyConfiguration -identity $useralias
    return $MailboxARC
}

function GET-ARCSTATE {
    $MailboxAR = get-arc
	Write-Host "Current Auto Reply State is :" + $MailboxARC.AutoReplyState
    return $MailboxARC.AutoReplyState
}

function Set-ARCSTATEScheduled {
    #if($MailboxARC -eq $undefinedVariable){
        $MailboxARC = get-arc
    #}
    #if($useralias -eq $undefinedVariable){
        $useralias = get-Alias
    #}
    #if($StartofShift -eq $undefinedVariable){
        $StartofShift = GetShiftTime("start")
    #}
    #if($EndofShift -eq $undefinedVariable){
        $EndofShift = GetShiftTime("end")
    #}
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
    Set-MailboxAutoReplyConfiguration –identity $useralias `
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

    $MessageFilePath = "C:\Users\" + $CurrentUser + "\OneDrive - Microsoft\Desktop\oof message script\OOF Message"

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
	$MessageFilePath = "C:\Users\" + $CurrentUser + "\OneDrive - Microsoft\Desktop\oof message script\OOF Message"
	if(($MessageFilePath + ".txt")){
		Write-Host "Setting the same OOF Message for both Internal and External"
		$Message = [System.IO.File]::ReadAllText($MessageFilePath+".txt")
		Write-Output $Message
		Set-MailboxAutoReplyConfiguration –identity $CUAlias –ExternalMessage $Message -InternalMessage $Message
	}
    else{  
		Write-Host "Differenet External and Internal Messages"
		$MessageFilePath = $MessageFilePath + "_External.txt"
		$Message = [System.IO.File]::ReadAllText($MessageFilePath)
		Write-Output ("Setting External Message`n`n" + $Message)
		Set-MailboxAutoReplyConfiguration –identity $CUAlias –ExternalMessage $Message
		$MessageFilePath = $MessageFilePath - "External.txt" + "Internal.txt"
		$Message = [System.IO.File]::ReadAllText($MessageFilePath)
		Write-Output ("Setting Internal Message`n`n" + $Message)
		Set-MailboxAutoReplyConfiguration –identity $CUAlias -InternalMessage $Message
    }
}

function GetShiftTime($startorend) { 
    $prompttext = "Enter when you " + $startorend + " your work day"
	$shifttime = Read-Host -Prompt $prompttext
	$shifttimeout = [datetime] $shifttime
    return $shfittimeout
}

function DisconnectEXO {
    Disconnect-ExchangeOnline -Confirm:$false
}
