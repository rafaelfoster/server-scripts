On Error Resume Next
' --------------------------------------------------------------------------------------------
' Criação do Objeto para criação/gravação do arquivo
set WshShell     = CreateObject("WScript.Shell")
Set objRegEx     = CreateObject("VBScript.RegExp")
Set WshNetwork   = CreateObject("WScript.Network")
Set objFSO       = CreateObject("Scripting.FileSystemObject")
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")

Const HKEY_LOCAL_MACHINE = &H80000002
Const ForReading = 1
Const ForWriting = 2
Const ForAppend  = 8
Const OverwriteExisting = TRUE

' --------------------------------------------------------------------------------------------
' Definição de Variaveis
strComputer = "."
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
strSessionName = wshShell.ExpandEnvironmentStrings( "%SESSIONNAME%" )
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
Log_Symantec = "\\rodrimar.com.br\Ti\AD-MGT\Logs\Log_Symantec\Log_Symantec.csv" 

	if ( inStr(LCase(strSessionName),"rdp") <> 0 OR inStr(LCase(strComputerName),"ctx") <> 0  OR inStr(LCase(strComputerName),"rod-") <> 0 ) Then
		Wscript.Quit
	End If

' Aguardar 3 minutos antes de iniciar
'Wscript.Sleep 180000

' -------------------------------------------------
' Definição do Regex 
objRegEx.Global = True
objRegEx.IgnoreCase = True
objRegEx.Pattern = "(" & strComputerName & ")\;(.+)\;(.+)\;(.+)\;(.+)"

' --------------------------------------------------------------------------------------------------------------------------------
' Logs de Auditoria de usuários

If objFSO.FileExists(Log_Symantec) = False Then
	Set objCriaLog = objFSO.CreateTextFile(Log_Symantec)
	objCriaLog.WriteLine "ComputerName;UserName;IPAddress;CurrentGroup;SymantecRegKey"
	objCriaLog.WriteLine
	objCriaLog.Close
End If

strKeyStatus = "HKLM\SOFTWARE\Symantec\Symantec Endpoint Protection\SMC\SYLINK\SyLink\CommunicationStatus"
SymantecStatus = WshShell.RegRead(strKeyStatus)

If IsEmpty(SymantecStatus) Then
	strKeyStatus = "HKLM\SOFTWARE\Symantec\Symantec Endpoint Protection\SMC\SYLINK\SyLink\LastServer"
	SymantecStatus = WshShell.RegRead(strKeyStatus)

	If IsEmpty(SymantecStatus) Then
		SymantecStatus = "Key empty or not found!"
	End If
	
End If
	
strKeyGroup = "HKLM\SOFTWARE\Symantec\Symantec Endpoint Protection\SMC\SYLINK\SyLink\CurrentGroup"
SymantecGroup = WshShell.RegRead(strKeyGroup)

strData = strComputerName &  ";" & strUserName & ";" & GetNetworkInformation & ";" & SymantecGroup & ";" & SymantecStatus

	Set objReadLog = objFSO.OpenTextFile(Log_Symantec, ForReading, True)
	strFileContent = objReadLog.ReadAll

	if ( inStr(LCase(strFileContent), LCase(strComputerName) ) <> 0 ) Then
			
			strFileContentEdited = objRegEx.Replace( strFileContent, strData )
			objReadLog.Close

		Set objCriaLog = objFSO.OpenTextFile(Log_Symantec, ForWriting, True)
			objCriaLog.Write strFileContentEdited 

	Else

		Set objCriaLog = objFSO.OpenTextFile(Log_Symantec, ForAppend, True)
			objCriaLog.WriteLine strData

	End If

objCriaLog.Close
objReadLog.Close

Function GetNetworkInformation()

	' List IP Configuration Data

	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

	Set colAdapters = objWMIService.ExecQuery _
		("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
	 
	n = 1
	 
	For Each objAdapter in colAdapters
	
	   If Not IsNull(objAdapter.IPAddress) Then
		  For i = 0 To UBound(objAdapter.IPAddress)
			 NetworkInformation = objAdapter.IPAddress(0)
		  Next
	   End If
	   n = n + 1

   Next

	GetNetworkInformation = NetworkInformation

End Function