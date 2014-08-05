On Error Resume Next
' --------------------------------------------------------------------------------------------
' Criação do Objeto para criação/gravação do arquivo
Set args             = WScript.Arguments
Set WshShell         = CreateObject("WScript.Shell")
Set WshNetwork       = CreateObject("WScript.Network")
Set objFSO           = CreateObject("Scripting.FileSystemObject")
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
Set SystemSet        = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 

Const ForReading = 1
Const ForWriting = 2
Const ForAppend  = 8
Const OverwriteExisting = TRUE

' --------------------------------------------------------------------------------------------
' Definição de Variaveis
strComputer = "."
strLogonType = args.Item(0)
strIPAddress = GetNetworkInformation
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
strUserPath = wshShell.ExpandEnvironmentStrings( "%USERPROFILE%" )
strLogonServer = wshShell.ExpandEnvironmentStrings( "%LOGONSERVER%" )
strSessionName = wshShell.ExpandEnvironmentStrings( "%SESSIONNAME%" )
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
strMonthName = getMonth 

Log_Auditoria_User = "\\SERVIDOR\Log_Auditoria\Log_Auditoria_Users\" 
Log_Auditoria_Workstation = "\\SERVIDOR\Log_Auditoria\Log_Auditoria_Workstations\"

Log_Header = "Data: " & date & " - " & time

' Aguardar 1 minuto antes de iniciar
'Wscript.Sleep 60000

'------------ Creating Year folders
If objFSO.FolderExists(Log_Auditoria_User & "\" & Year(date)) = False Then
	objFSO.CreateFolder(Log_Auditoria_User & "\" & Year(date))
End If
If objFSO.FolderExists(Log_Auditoria_User & "\" & Year(date) & "\" & strMonthName ) = False Then
	objFSO.CreateFolder(Log_Auditoria_User & "\" & Year(date) & "\" & strMonthName )
End If

'------------ Creating month folders
If objFSO.FolderExists(Log_Auditoria_Workstation & "\" & Year(date)) = False Then
	objFSO.CreateFolder(Log_Auditoria_Workstation & "\" & Year(date))
End If
If objFSO.FolderExists(Log_Auditoria_Workstation & "\" & Year(date) & "\" & strMonthName ) = False Then
	objFSO.CreateFolder(Log_Auditoria_Workstation & "\" & Year(date) & "\" & strMonthName )
End If

Log_Auditoria_User = Log_Auditoria_User & "\" & Year(date) & "\" & strMonthName & "\Log_Audit_" & strUserName & ".txt" 
Log_Auditoria_Workstation =  Log_Auditoria_Workstation & "\" & Year(date) & "\" & strMonthName & "\Log_Audit_" & strComputerName & ".txt" 


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

objCriaLog.WriteLine "Data: " & date & " - " & time & " - Computador: " & strComputerName & " - Tipo: " & strLogonType & vbTab & "- IP: " & strIPAddress
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

objCriaLog.WriteLine "Data: " & date & " - " & time & " - Usuario: " & strUserName & " - Tipo: " & strLogonType & vbTab & "- IP: " & strIPAddress
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
					if IsEmpty(NetworkInformation) Then
						NetworkInformation = objAdapter.IPAddress(i)
					Else
						NetworkInformation = NetworkInformation & ", " & objAdapter.IPAddress(i)
					End If
				End If
		  Next
	   End If
	   n = n + 1

   Next	

   GetNetworkInformation = NetworkInformation

End Function

function getMonth()
	strMonth = MonthName(Month(Date))
	Select Case LCase(strMonth)
		Case "january"
			strThisMonth = "janeiro"
		Case "february"
			strThisMonth = "fevereiro"
		Case "march"
			strThisMonth = "marco"
		Case "april"
			strThisMonth = "abril"
		Case "may"
			strThisMonth = "maio"
		Case "june"
			strThisMonth = "junho"
		Case "july"
			strThisMonth = "julho"
		Case "august"
			strThisMonth = "agosto"
		Case "september"
			strThisMonth = "setembro"
		Case "october"
			strThisMonth = "outubro"
		Case "november"
			strThisMonth = "novembro"
		Case "december"
			strThisMonth = "dezembro"
		Case Else
			strThisMonth = strMonth
	End Select
	getMonth = strThisMonth
End Function