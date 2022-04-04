. "./OOFFunctions.ps1" #include all the fancy functions
Pause
get-Alias #get username and suffix default is username from local machine plus @microsoftsupport.com
Pause
ConnectAlias2EXO # connect to Exchange online
Pause
get-ARC	#check for local config, if none get auto reply config, use current message
Pause
DisconnectEXO