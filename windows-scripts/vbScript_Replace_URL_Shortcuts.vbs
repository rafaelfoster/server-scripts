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
objRegEx.Pattern = "(\s?\w+):\/\/(\s?[\w@][\w.:@]+)\/?([\w\.?=%&=\-@/$,]*)" ' Definição de REGEX para URLS Ex.: http://www.servidor.com.br/pasta/arquivo.html

strComputer = "."
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
strUserPath = wshShell.ExpandEnvironmentStrings( "%USERPROFILE%" )
strSessionName = wshShell.ExpandEnvironmentStrings( "%SESSIONNAME%" )
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )

Log_Folder = "\\SERVIDOR\COMPARTILHAMENTO\Log"
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

	if ( inStr(LCase(objFile.Name),".url") <> 0 ) Then

		Set objArquivo = objFSO.OpenTextFile(objStartFolder & "\" & objFile.Name, ForReading, True)
		strFileContent = objArquivo.ReadAll

		if ( inStr(LCase(strFileContent), "STRING SERVIDOR ANTIGO" ) <> 0 ) Then
			NumberFiles = NumberFiles + 1
	
			If Not objFSO.FolderExists(Log_Folder & "\Shortcuts\" & strUserName) Then
				objFSO.CreateFolder(Log_Folder & "\Shortcuts\" & strUserName)
			End If

			objFSO.CopyFile objStartFolder & "\" & objFile.Name, Log_Folder & "\Shortcuts\" & strUserName & "\"
			
			strFileContentEdited = objRegEx.Replace( strFileContent, "http://www.servidor.com.br" )

			Set objArquivoNovo = objFSO.OpenTextFile(objStartFolder & "\" & objFile.Name, ForWriting, True)
			objArquivoNovo.WriteLine strFileContentEdited

			objArquivoNovo.Close

			objCriaLog.WriteLine "Alterando o arquivo: " & objFile.Name

		End If

		objArquivo.Close

	End If

Next

objCriaLog.Close

If NumberFiles = 0 Then
	If Not objFSO.FolderExists(Log_Folder & "\Shortcuts\" & strUserName) Then
		objFSO.DeleteFile Log_File, True
	End If
End If 