On Error Resume Next
' --------------------------------------------------------------------------------------------
' Criação do Objeto para criação/gravação do arquivo
Set WshShell     = CreateObject("WScript.Shell")
Set WshNetwork   = CreateObject("WScript.Network")
Set objFSO       = CreateObject("Scripting.FileSystemObject")
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
Set SystemSet    = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 

Const ForReading = 1
Const ForWriting = 2
Const ForAppend  = 8
Const OverwriteExisting = TRUE

' --------------------------------------------------------------------------------------------
' Definição de Variaveis
strComputer = "."
strIPAddress = GetNetworkInformation
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
strUserPath = wshShell.ExpandEnvironmentStrings( "%USERPROFILE%" )
strLogonServer = wshShell.ExpandEnvironmentStrings( "%LOGONSERVER%" )
strSessionName = wshShell.ExpandEnvironmentStrings( "%SESSIONNAME%" )
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )

strLogsBasePath = "\\SERVIDOR\COMPARTILHAMENTO"
Log_Auditoria_User = strLogsBasePath & "\Log_Audit_" & strUserName & ".txt" 
Log_Auditoria_Workstation = strLogsBasePath & "\Log_Audit_" & strComputerName & ".txt" 

' Aguardar 1 minuto antes de iniciar
Wscript.Sleep 60000

' --------------------------------------------------------------------------------------------------------------------------------
' Logs de Auditoria de usuários

If objFSO.FileExists(Log_Auditoria_User) Then
	Set objCriaLog = objFSO.OpenTextFile(Log_Auditoria_User, ForAppend, TRUE)
Else
	Set objCriaLog = objFSO.CreateTextFile(Log_Auditoria_User)
	objCriaLog.WriteLine "---------------------------------------------------------------------------------------"
	objCriaLog.WriteLine "REGISTRO DE AUDITORIA - USUARIO: " & strUserName
	objCriaLog.WriteLine
End If

objCriaLog.WriteLine "Data: " & date & " - " & time & " - Computador: " & strComputerName & " - IP: " & strIPAddress
objCriaLog.Close

' --------------------------------------------------------------------------------------------------------------------------------
' Logs de Auditoria de Computadores

If objFSO.FileExists(Log_Auditoria_Workstation) Then
	Set objCriaLog = objFSO.OpenTextFile(Log_Auditoria_Workstation, ForAppend, TRUE)
Else
	Set objCriaLog = objFSO.CreateTextFile(Log_Auditoria_Workstation)
	objCriaLog.WriteLine "---------------------------------------------------------------------------------------"
	objCriaLog.WriteLine "REGISTRO DE AUDITORIA - COMPUTADOR: " & strComputerName
	objCriaLog.WriteLine
End If

objCriaLog.WriteLine "Data: " & date & " - " & time & " - Usuario: " & strUserName & " - IP: " & strIPAddress
objCriaLog.Close

Function GetNetworkInformation()

	' List IP Configuration Data

	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

	Set colAdapters = objWMIService.ExecQuery _
		("SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
	 
	n = 1
	 
	For Each objAdapter in colAdapters
	
	   If Not IsNull(objAdapter.IPAddress) Then
		  For i = 0 To UBound(objAdapter.IPAddress)
				If inStr(objAdapter.IPAddress(i),".") Then
					NetworkInformation = objAdapter.IPAddress(i)
				End If
		  Next
	   End If
	   n = n + 1

   Next

	GetNetworkInformation = NetworkInformation

End Function
