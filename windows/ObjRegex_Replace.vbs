On Error Resume Next
' --------------------------------------------------------------------------------------------
' Criação do Objeto para criação/gravação do arquivo
set WshShell     = CreateObject("WScript.Shell")
Set objRegEx     = CreateObject("VBScript.RegExp")
Set WshNetwork   = CreateObject("WScript.Network")
Set objFSO       = CreateObject("Scripting.FileSystemObject")

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

' -------------------------------------------------
' Definição do Regex 
objRegEx.Global = True
objRegEx.IgnoreCase = True
objRegEx.Pattern = "(.+)\;(.+)\;(.+)\;(.+)\;(.+)" 'Definição do string de busca do Regex

Set objReadLog = objFSO.OpenTextFile("ARQUIVO_DE_TEXTO", ForReading, True)
strFileContent = objReadLog.ReadAll

	if ( inStr(LCase(strFileContent), LCase(strComputerName) ) <> 0 ) Then
			
			strFileContentEdited = objRegEx.Replace( strFileContent, "TEXTO NOVO" )
			objReadLog.Close

		' Escreve o novo arquivo com o conteudo novo
		Set objCriaLog = objFSO.OpenTextFile("ARQUIVO_DE_TEXTO", ForWriting, True)
			objCriaLog.Write strFileContentEdited 

	End If

objCriaLog.Close
objReadLog.Close