'
' Script: vbs_MigrateOutlookExchangeProfile.vbs
'
' This script change the registry keys to 
' disable an old Exchange Server and enable a New Exchange Server 
' pointed in Microsoft Outlook and generate a log file.
' 
' Running it from the command line with the parameter DEBUG, (as showed above)
' the script will return the informations to the console at the same time 
' the log will be generated.
'
' Script created by Rafael Foster (rafaelgfoster at gmail dot com)
'

ON ERROR RESUME NEXT 

CONST OldServer = "ROD-MAIL01"
CONST NewServer = "POMBO"
CONST ServerDomain = "rodrimar.com.br"
'----------------------------
CONST FQDNValue = "001e6608"
CONST ProfileKey = "001e6750"
CONST FQDNBinary = "001f662a"
CONST NetBiosValue = "001e6602"
CONST xFiveHundredValue = "001e6612"
Const HKEY_CURRENT_USER = &H80000001
Const Log_Anexar = 2 '( 1 = Read, 2 = Write, 8 = Append )

computerName = "."
Set args      = WScript.Arguments
Set WshShell  = CreateObject("WScript.Shell")
Set objFSO    = CreateObject("Scripting.FileSystemObject")
Set SystemSet = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 
set objReg    = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & computerName & "\root\default:StdRegProv")

strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
strUserPath = wshShell.ExpandEnvironmentStrings( "%USERPROFILE%" )
strLogonServer = wshShell.ExpandEnvironmentStrings( "%LOGONSERVER%" )
strSessionName = wshShell.ExpandEnvironmentStrings( "%SESSIONNAME%" )
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
strProgFiles = wshShell.ExpandEnvironmentStrings( "%PROGRAMFILES%" )
strProgFilesx86 = wshShell.ExpandEnvironmentStrings( "%PROGRAMFILES(x86)%" )
Log_File = "\\rodrimar.com.br\Ti\AD-MGT\Logs\Log_MigrateExchange\Log_MigrateExchange_" & strUserName & ".txt" 

For each System in SystemSet 
	SysVersion     = System.Caption & " SP" & System.ServicePackMajorVersion & " " & System.BuildNumber
Next
if ( inStr(LCase(SysVersion),"server") <> 0 ) Then
	Wscript.Quit
End If

If objFSO.FileExists(Log_File) Then
	Set objCriaLog = objFSO.OpenTextFile(Log_File, Log_Anexar, True)
Else
	Set objCriaLog = objFSO.CreateTextFile(Log_File)
End If

strOfficeVersion = GetOfficeVersionNumber() 
Log_Header = "Data: " & date & " - " & time
WriteOutput "-------------------------------------------------------------------------------------------------------"
WriteOutput Log_Header
WriteOutput ""
WriteOutput "-------[ System Information ]---------------------------------------------------"
WriteOutput vbCr
WriteOutput "Usuario               : " & strUserName
WriteOutput "Estacao de Trabalho   : " & strComputerName
WriteOutput "Servidor de Logon     : " & strLogonServer
WriteOutput "Sistema Operacional   : " & SysVersion
WriteOutput "Versao do Office      : " & strOfficeVersion
WriteOutput vbCr

FqdnBinaryValue = "50,00,4f,00,4d,00,42,00,4f,00,2e,00,52,00,4f,00,44,00,52,00,49,00,4d,00,41,00,52,00,2e,00,43,00,4f,00,4d,00,2e,00,42,00,52,00,00,00"

if strOfficeVersion < 15 then
	BASE_KEY = "Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"
Else
	BASE_KEY = "Software\Microsoft\Office\15.0\Outlook\Profiles"
End If

objReg.EnumKey HKEY_CURRENT_USER, BASE_KEY, arrSubKeys
 
for each subkey in arrSubKeys
	strProfile = subkey
next

	WriteOutput vbCr
	WriteOutput "Profile: " & strProfile
	WriteOutput "Registry Path: " & BASE_KEY
	strSubKey = FindProfileKey(HKEY_CURRENT_USER, BASE_KEY)
	
set objReg = nothing

sub gethashkey(hash)

	subKeyPath = BASE_KEY & "\" & strProfile & "\" & hash

	objReg.GetStringValue HKEY_CURRENT_USER, subKeyPath, NetBiosValue, nbServerName
	objReg.GetStringValue HKEY_CURRENT_USER, subKeyPath, FQDNValue, fqdn
	objReg.GetStringValue HKEY_CURRENT_USER, subKeyPath, xFiveHundredValue, X500

	fqdn = ucase(fqdn)

	WriteOutput vbCr
	WriteOutput "Current Values (" & hash & ")"
	WriteOutput "--------------"
	WriteOutput "Netbios Value: " & nbServerName
	WriteOutput "FQDN Value: " & lcase(fqdn)
	WriteOutput "X500 Value: " & X500
	WriteOutput "FQDN Binary Value: " & fqdnbin

	if ( instr(nbServerName, NewServer) = 0 OR instr(fqdn, NewServer) = 0 OR instr(X500, NewServer) = 0 ) then

		nbNewServerName = Replace(ucase(nbServerName), nbServerName, NewServer)
		fqdn = lcase(NewServer) & "." & ServerDomain
		x500 = Replace(X500, nbServerName, NewServer)

		WriteOutput vbCr
		WriteOutput "New Values"
		WriteOutput "--------------"
		WriteOutput "Netbios Value: " & nbNewServerName
		WriteOutput "FQDN Value: " & fqdn 
		WriteOutput "X500 Value: " & X500
		WriteOutput "FQDN Binary Value: " & FqdnBinaryValue

		objReg.SetStringValue HKEY_CURRENT_USER, subKeyPath, NetBiosValue, nbNewServerName
		objReg.SetStringValue HKEY_CURRENT_USER, subKeyPath, FQDNValue, fqdn
		objReg.SetStringValue HKEY_CURRENT_USER, subKeyPath, xFiveHundredValue, X500
		WshShell.RegDelete "HKEY_CURRENT_USER\" & subKeyPath & "\" & FQDNBinary

	end if
	
End sub

Function GetOfficeVersionNumber()
    GetOfficeVersionNumber = ""
    Dim sTempValue
	sTempValue = WshShell.RegRead("HKCR\Excel.Application\CurVer\")
    If Len(sTempValue) > 2 Then GetOfficeVersionNumber = Replace(Right(sTempValue, 2), ".", "") 
End Function    

function FindProfileKey(hive, key)
	Set reg = GetObject("winmgmts://./root/default:StdRegProv")
	
	reg.EnumKey hive, key, arrSubKeys1
	If Not IsNull(arrSubKeys1) Then
		For Each strSubkey2 In arrSubKeys1
			FindProfileKey hive, key & "\" & strSubkey2
		Next
	End If

	arrKey = split(key,"\",-1,1)
	For each strregSubkey in arrKey
		strSubkeyHash = strregSubkey
	Next

	if reg.enumValues( HKEY_CURRENT_USER, key, valueNames, types ) = 0 then
		if isArray( valueNames ) then
			for i = 0 to UBound( valueNames )
				reg.getStringValue HKEY_CURRENT_USER, key, valueNames(i), value
				if ( valueNames(i) = ProfileKey ) then
					gethashkey strSubkeyHash
				End If
			next
		end if
	end if
End Function

function WriteOutput(strText)
	if (args.Count > 0) Then
		if (inStr(LCase(args.Item(0)),"debug") <> 0) Then
			Wscript.Echo strText
		End If
	End If
	objCriaLog.WriteLine strText
End Function