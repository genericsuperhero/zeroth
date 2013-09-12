Option Explicit
 
const HKEY_LOCAL_MACHINE = &H80000002
 
dim ProductName, ProductKey
 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
sub GetSymantecProductKey()
 
dim oReg, sPath, aKeys, sName, sKey
Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
 
sPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
oReg.EnumKey HKEY_LOCAL_MACHINE, sPath, aKeys
 
For Each sKey in aKeys
	oReg.GetStringValue HKEY_LOCAL_MACHINE, sPath & "\" & sKey, "DisplayName", sName
	If Not IsNull(sName) Then 
		if (sName = "Symantec Endpoint Protection") then
			ProductKey = sKey
			ProductName = sName
		end if
	end if
Next
 
end sub
 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
sub RemoveSymantec(key, name)
 
dim cmd, objShell, iReturn
cmd = "C:\windows\system32\msiexec.exe /q/x " & key
 
set objShell = wscript.createObject("wscript.shell")
 
objShell.LogEvent 0, "Removing the program [" & name & "] under Product Key [" & key & "]" & vbCrLf & "Executing command: " & vbCrLf & cmd
 
iReturn=objShell.Run(cmd,1,TRUE)
 
if (iReturn = 0) then
	objShell.LogEvent 0, "Program [" & name & "] was successfully removed"
else
	objShell.LogEvent 0, "Failed to remove the program [" & name & "]."
end if
 
Set objShell = Nothing  
 
end sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

ProductKey = ""
ProductName = ""
 
call GetSymantecProductKey()
if Not (ProductKey = "") then
	call RemoveSymantec(ProductKey, ProductName)
end if
