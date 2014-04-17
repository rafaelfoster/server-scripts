On Error Resume Next
Set WshShell = CreateObject("WScript.Shell")
Set WshNetwork = CreateObject("WScript.Network")
Set objFSO = CreateObject("Scripting.FileSystemObject")

NumberFiles = 0
Const ForReading = 1
Const ForWriting = 2
Const ForAppend  = 8
Const OverwriteExisting = TRUE

strComputer = "."
strPrintServer = "rod-printsb"
strAddImpressoraAsDefault = TRUE
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
strSessionName = wshShell.ExpandEnvironmentStrings( "%SESSIONNAME%" )
strUserPath = wshShell.ExpandEnvironmentStrings( "%USERPROFILE%" )
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
Log_Folder = "\\rodrimar.com.br\TI\AD-MGT\LOGS\Log_Script_impressoras_SegurancaSB"
Log_File = Log_Folder & "\Log_Impressora_" & strUserName & ".txt"

	if ( inStr(LCase(strSessionName),"rdp") <> 0 ) Then
		Wscript.Quit
	End If

'Aguardar 1 minuto antes de iniciar
Wscript.Sleep 60000

if (inStr(LCase(strUserName),"ebsantos")    <> 0 OR _
	inStr(LCase(strUserName),"vosantos")    <> 0 OR _
	inStr(LCase(strUserName),"crogerio")    <> 0 OR _
	inStr(LCase(strUserName),"aroliveira")  <> 0 OR _
	inStr(LCase(strUserName),"albano")      <> 0 OR _
	inStr(LCase(strUserName),"rpereira")    <> 0 OR _
	inStr(LCase(strUserName),"bbarrios")    <> 0 OR _
	inStr(LCase(strUserName),"dsouza")      <> 0 OR _
	inStr(LCase(strUserName),"cssilva")     <> 0 OR _
	inStr(LCase(strUserName),"cspinho")     <> 0 OR _
	inStr(LCase(strUserName),"cspinho")     <> 0 ) Then

	If objFSO.FileExists(Log_File) Then
		Set objCriaLog = objFSO.OpenTextFile(Log_File, ForAppend, True)
	Else
		Set objCriaLog = objFSO.CreateTextFile(Log_File)
	End If

	objCriaLog.WriteLine 
	objCriaLog.WriteLine "Data: " & date & " - " & time & " - Computador: " & strComputerName & " | Usuario: " & strUserName

	strAddImpressora = "SEG_COLOR"
	
	' Adiciona a Impressora desejada
	WshNetwork.AddWindowsPrinterConnection "\\" & strPrintServer & "\" & strAddImpressora

	If strAddImpressoraAsDefault = True Then
		WshNetwork.SetDefaultPrinter "\\" & strPrintServer & "\" & strAddImpressora
	End If

	Set objWMIService = GetObject("winmgmts:" _ 
		& "{impersonationLevel=impersonate}!\\" & strComputerName & "\root\cimv2") 

	Set colInstalledPrinters = objWMIService.ExecQuery _ 
		("Select * from Win32_Printer") 
		
	For Each objPrinter in colInstalledPrinters
		if ( InStr(LCase(objPrinter.Name), LCase(strAddImpressora) ) <> 0 ) Then
			objCriaLog.WriteLine "Impressora Instalada: " &  objPrinter.name
		End IF
	Next

	Wscript.Quit
	
End If

if (inStr(LCase(strUserName),"wxavier")     <> 0 OR _
	inStr(LCase(strUserName),"wcabral")     <> 0 OR _
	inStr(LCase(strUserName),"jmedeiros")   <> 0 OR _
	inStr(LCase(strUserName),"segurancasb") <> 0 OR _
	inStr(LCase(strUserName),"ascarvalho")  <> 0 OR _
	inStr(LCase(strUserName),"cspinho")     <> 0 ) Then

	If objFSO.FileExists(Log_File) Then
		Set objCriaLog = objFSO.OpenTextFile(Log_File, ForAppend, True)
	Else
		Set objCriaLog = objFSO.CreateTextFile(Log_File)
	End If

	objCriaLog.WriteLine 
	objCriaLog.WriteLine "Data: " & date & " - " & time & " - Computador: " & strComputerName & " | Usuario: " & strUserName

	strAddImpressora = "SEG_SB"
 
	' Adiciona a Impressora desejada
	WshNetwork.AddWindowsPrinterConnection "\\" & strPrintServer & "\" & strAddImpressora

	If strAddImpressoraAsDefault = True Then
		WshNetwork.SetDefaultPrinter "\\" & strPrintServer & "\" & strAddImpressora
	End If

	Set objWMIService = GetObject("winmgmts:" _ 
		& "{impersonationLevel=impersonate}!\\" & strComputerName & "\root\cimv2") 

	Set colInstalledPrinters = objWMIService.ExecQuery _ 
		("Select * from Win32_Printer") 
		
	For Each objPrinter in colInstalledPrinters
		if ( InStr(LCase(objPrinter.Name), LCase(strAddImpressora) ) <> 0 ) Then
			objCriaLog.WriteLine "Impressora Instalada: " &  objPrinter.name
		End IF
	Next

	Wscript.Quit
	
End If