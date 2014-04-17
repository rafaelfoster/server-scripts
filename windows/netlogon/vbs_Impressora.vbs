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
strAddImpressora = "PORTO_01"
strAddImpressoraAsDefault = TRUE
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
strSessionName = wshShell.ExpandEnvironmentStrings( "%SESSIONNAME%" )
strUserPath = wshShell.ExpandEnvironmentStrings( "%USERPROFILE%" )
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
Log_Folder = "\\rodrimar.com.br\TI\AD-MGT\LOGS\Log_Script_impressoras_OperPort"
Log_File = Log_Folder & "\Log_Impressora_" & strUserName & ".txt"

	if ( inStr(LCase(strSessionName),"rdp") <> 0 OR inStr(LCase(strUserName),"ahermida") <> 0 OR inStr(LCase(strUserName),"lrocha") <> 0 OR inStr(LCase(strUserName),"fgoncalves") <> 0 OR inStr(LCase(strUserName),"aguirre") <> 0 OR inStr(LCase(strUserName),"ehenriques") <> 0 OR inStr(LCase(strUserName),"salves") <> 0 ) Then
		Wscript.Quit
	End If

'Aguardar 1 minuto antes de iniciar
'Wscript.Sleep 60000

If objFSO.FileExists(Log_File) Then
	Set objCriaLog = objFSO.OpenTextFile(Log_File, ForAppend, True)
Else
	Set objCriaLog = objFSO.CreateTextFile(Log_File)
End If

objCriaLog.WriteLine 
objCriaLog.WriteLine "Data: " & date & " - " & time & " - Computador: " & strComputerName & " | Usuario: " & strUserName

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