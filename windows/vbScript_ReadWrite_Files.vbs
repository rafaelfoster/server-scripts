'On Error Resume Next
' --------------------------------------------------------------------------------------------
' Criação do Objeto para criação/gravação do arquivo
set WshShell     = CreateObject("WScript.Shell")
Set objFSO       = CreateObject("Scripting.FileSystemObject")

Const ForReading = 1
Const ForWriting = 2
Const ForAppend  = 8
Const OverwriteExisting = TRUE

strArquivo = "C:\Windows\Temp\texto.txt"

If objFSO.FileExists(strArquivo) Then
	Set objCriaLog = objFSO.OpenTextFile(strArquivo, ForAppend, TRUE)
Else
	Set objCriaLog = objFSO.CreateTextFile(strArquivo)
End If

objCriaLog.WriteLine "Escreve Linha COM \n no final"
objCriaLog.Write     "Escreve Linha SEM \n no final"
objCriaLog.Close

set objCriaLog = objFSO.OpenTextFile(strArquivo, ForReading, TRUE)
strFileContent = objCriaLog.ReadAll
Wscript.Echo strFileContent


objCriaLog.Close
