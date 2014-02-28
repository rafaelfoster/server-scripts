' For URL shortcuts, only ".FullName" and ".TargetPath" are valid
'WScript.Echo "Full Name         : " & objShortcut.FullName
'WScript.Echo "Arguments         : " & objShortcut.Arguments
'WScript.Echo "Working Directory : " & objShortcut.WorkingDirectory
'WScript.Echo "Target Path       : " & objShortcut.TargetPath
'WScript.Echo "Icon Location     : " & objShortcut.IconLocation
'WScript.Echo "Hotkey            : " & objShortcut.Hotkey
'WScript.Echo "Window Style      : " & objShortcut.WindowStyle
'WScript.Echo "Description       : " & objShortcut.Description

On Error Resume Next
Set WshShell = CreateObject("WScript.Shell")
Set objRegEx = CreateObject("VBScript.RegExp")
Set objFSO = CreateObject("Scripting.FileSystemObject")

NumberFiles = 0
Const ForReading = 1
Const ForWriting = 2
Const ForAppend  = 8
Const OverwriteExisting = TRUE

' -------------------------------------------------
' Definição do Regex 
objRegEx.Global = True
objRegEx.IgnoreCase = True
objRegEx.Pattern = "^(\\\\((servidor1)|(servidor2)|(192.168.20.(\d+)))\\pasta[s]?\\)"

strComputer = "."
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
strUserPath = wshShell.ExpandEnvironmentStrings( "%USERPROFILE%" )
strSessionName = wshShell.ExpandEnvironmentStrings( "%SESSIONNAME%" )
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )

Log_Folder = "\\SERVIDOR\compartilhamento\Logs"
Log_File = Log_Folder & "\Log_Shortcut_" & strUserName & ".txt"

	if ( inStr(LCase(strSessionName),"rdp") <> 0 ) Then
		Wscript.Quit
	End If

'Aguardar 1 minuto antes de iniciar
Wscript.Sleep 60000

' ------------------------------------------------------------------------------------------------------------------------------------------------
' Inicia a geração do Arquivo de Log
If objFSO.FileExists(Log_File) Then
	Set objCriaLog = objFSO.OpenTextFile(Log_File, ForAppend, True)
Else
	Set objCriaLog = objFSO.CreateTextFile(Log_File)
End If

objCriaLog.WriteLine 
objCriaLog.WriteLine "Data: " & date & " - " & time & " - Computador: " & strComputerName & " | Usuario: " & strUserName

objStartFolder = strUserPath & "\Desktop\"

Set objFolder = objFSO.GetFolder(objStartFolder)
Set colFiles = objFolder.Files

For Each objFile in colFiles

	if ( inStr(LCase(objFile.Name),".lnk") <> 0 ) Then
	'----------------------------

		Set objShortcut = wshShell.CreateShortcut(objFile.Path)

		TargetPath = objShortcut.TargetPath
		WorkingDirectory = objShortcut.WorkingDirectory
		
		' Se o destino do atalho for para Destination_Folder, executa a troca
		if ( inStr(LCase(objShortcut.TargetPath), "\Destination_Folder" ) <> 0  ) Then
			NumberFiles = NumberFiles + 1

			'WScript.Echo "Full Name         :        " & objShortcut.FullName & vbCR _
			'		   & "Target Path       :        " & TargetPath
			
			If Not objFSO.FolderExists(Log_Folder & "\Shortcuts\" & strUserName) Then
				objFSO.CreateFolder(Log_Folder & "\Shortcuts\" & strUserName)
			End If

			objFSO.CopyFile objShortcut.FullName, Log_Folder & "\Shortcuts\" & strUserName & "\"
			
			'Wscript.Echo TargetPath & vbCR & vbCR & objRegEx.Replace( TargetPath, "\\servidor\compartilhamento_Novo\"  )
			
			objShortcut.TargetPath = objRegEx.Replace( TargetPath , "\\servidor\compartilhamento_Novo\" )
			objShortcut.WorkingDirectory = objRegEx.Replace( WorkingDirectory , "\\servidor\compartilhamento_Novo\" )

			objShortcut.Save
			objCriaLog.WriteLine ""
			objCriaLog.WriteLine "Alterando o arquivo: " & objFile.Name
			objCriaLog.WriteLine "Caminho antigo: " &  TargetPath
			objCriaLog.WriteLine "Caminho novo: " & objRegEx.Replace( TargetPath , "\\servidor\compartilhamento_Novo\" )
			objCriaLog.WriteLine ""

		End If
		
		Set objShortcut = Nothing

	'-----------------------------

	End If

Next

objCriaLog.Close

If NumberFiles = 0 Then
	If Not objFSO.FolderExists(Log_Folder & "\Shortcuts\" & strUserName) Then
		objFSO.DeleteFile Log_File, True
	End If
End If