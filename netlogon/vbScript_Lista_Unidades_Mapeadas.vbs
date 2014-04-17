On Error Resume Next
' --------------------------------------------------------------------------------------------
' Criação do Objeto para criação/gravação do arquivo
Set WshShell     = CreateObject("WScript.Shell")
Set WshNetwork   = CreateObject("WScript.Network")
Set objFSO       = CreateObject("Scripting.FileSystemObject")
Set SystemSet    = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 
Const Log_Anexar = 2 '( 1 = Read, 2 = Write, 8 = Append )

' --------------------------------------------------------------------------------------------
' Definição de Variaveis
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
strUserPath = wshShell.ExpandEnvironmentStrings( "%USERPROFILE%" )
strLogonServer = wshShell.ExpandEnvironmentStrings( "%LOGONSERVER%" )
strSessionName = wshShell.ExpandEnvironmentStrings( "%SESSIONNAME%" )
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
Log_File = "\\rodrimar.com.br\Ti\AD-MGT\Logs\Log_Script_MappedDrives\Log_MappedDrives_" & strUserName & ".txt" 
Log_Shortcuts_File = "\\rodrimar.com.br\Ti\AD-MGT\Logs\Log_Script_DesktopShortcuts\Log_DesktopShortcuts_" & strUserName & ".txt"
Log_Header = "Data: " & date & " - " & time & " - Computador: " & strComputerName & " | Usuario: " & strUserName & " | Logado em: " & strLogonServer & " | Sessão: " & strSessionName & " | " 

	if ( inStr(LCase(strSessionName),"rdp") <> 0 OR inStr(LCase(strComputerName),"ctx") <> 0 ) Then
		Wscript.Quit
	End If

' Aguardar 3 minutos antes de iniciar
'Wscript.Sleep 180000

If objFSO.FileExists(Log_File) Then
	Set objCriaLog = objFSO.OpenTextFile(Log_File, Log_Anexar, True)
Else
	Set objCriaLog = objFSO.CreateTextFile(Log_File)
End If

For each System in SystemSet 
	SysOperation = System.Caption & " SP" & System.ServicePackMajorVersion & " " & System.BuildNumber
Next
Log_Header = Log_Header & SysOperation

objCriaLog.WriteLine 
objCriaLog.WriteLine Log_Header
objCriaLog.WriteLine 

strComputer = "."

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_MappedLogicalDisk")

objCriaLog.WriteLine "-------[ Unidades Mapeadas neste computador ]-----------------------------------"
For Each objItem in colItems
	objCriaLog.WriteLine "Unidade: " & objItem.Name & " - em: " & objItem.ProviderName
Next
objCriaLog.WriteLine "--------------------------------------------------------------------------------"
objCriaLog.WriteLine

' ----------------------------------------------------------------------------------------------------
' Coleta Shortcuts e URLs do usuário

If objFSO.FileExists(Log_Shortcuts_File) Then
	Set objCriaLog = objFSO.OpenTextFile(Log_Shortcuts_File, Log_Anexar, True)
Else
	Set objCriaLog = objFSO.CreateTextFile(Log_Shortcuts_File)
End If

objCriaLog.WriteLine
objCriaLog.WriteLine Log_Header
objCriaLog.WriteLine

objStartFolder = strUserPath & "\Desktop\"

Set objFolder = objFSO.GetFolder(objStartFolder)
Set colFiles = objFolder.Files

For Each objFile in colFiles
	
	strFileExt = objFSO.GetExtensionName(objFile.Name)
		
	'Wscript.Echo "Extensao: " & strFileExt 
	
	Select Case strFileExt
	
		Case "lnk"
			Set objShortcut = wshShell.CreateShortcut(objFile.Path)
			objCriaLog.WriteLine "Atalho:  " & objFile.Name

			if ( InStr(Lcase(SysOperation), "xp") <> 0 ) Then
				objCriaLog.WriteLine "Destino: " & objShortcut.TargetPath
			Else
				objCriaLog.WriteLine "Destino: " & objShortcut.TargetPath
			End If
			
			objCriaLog.WriteLine "----"
		
		Case "url"
			Set objShortcut = wshShell.CreateShortcut(objFile.Path)
			objCriaLog.WriteLine "URL:     " & objFile.Name
			
			if ( InStr(LCase(SysOperation), "xp") <> 0 ) Then
				objCriaLog.WriteLine "Destino: " & objShortcut.TargetPath
			Else
				objCriaLog.WriteLine "Destino: " & objShortcut.TargetPath
			End If
			
			objCriaLog.WriteLine "----"
	
	End Select
	
	Set objShortcut = Nothing

	'-----------------------------

Next

objCriaLog.WriteLine "--------------------------------------------------------------------------------"
objCriaLog.WriteLine

objCriaLog.Close
