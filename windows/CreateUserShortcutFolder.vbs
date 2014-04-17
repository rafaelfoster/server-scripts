On Error Resume Next
' ----------------------------------------------------------------------------------------------
' This script creates a folder called strDestFolderName as a SymLink in the Windows Sistems.
' This features allows to create a shortcut to a network share and the shortcut act as a Folder,
' not as a shortcut file.
'
' ----------------------------------------------------------------------------------------------
' Criação do Objeto para criação/gravação do arquivo
Set WshShell     = CreateObject("WScript.Shell")
Set objFSO       = CreateObject("Scripting.FileSystemObject")

Const ForReading = 1
Const ForWriting = 2
Const ForAppend  = 8
Const Attrib_System = 4
Const Attrib_ReadOnly = 1
Const OverwriteExisting = TRUE

intLoopControl = 0

While intLoopControl <> 6 
	
	Do
		strUserName = inputbox("Digite o nome do usuario")
	Loop Until Len(strUserName) <> 0

	Do
		strDepto = inputbox("Digite o nome do Depto")
	Loop Until Len(strDepto) <> 0

	Do
		strLocalidade = inputbox("Digite a Localidade")
	Loop Until Len(strLocalidade) <> 0

	strMsgConfirm = "Usuario: " & strUserName & vbCr _ 
				& "Depto:" & strDepto & vbCr _ 
				& "Local:" & strLocalidade & vbCr & vbCr _ 
				& "O caminho do MAIL do usuario sera:" & vbCr _ 
				& "\\SERVIDOR\Mail\" & strLocalidade & "\" & strUserName & vbCr _ 
				& "Confirma estes dados?"

	intLoopControl = MsgBox(strMsgConfirm,vbYesNo,"")

Wend

strTargetLink = "\\SERVIDOR\COMPARTILHAMENTO\" & strLocalidade & "\" & strUserName
strDestFolderPath  = "\\SERVIDOR\" & strDepto & "\Usuarios\" & strUserName
strDestFolderName = "PST"

'Criar Pasta " & strDestFolderName & " -> "\\SERVIDOR\" & strDepto & "\Usuarios\" & strUserName
Do
	Set objFolder = objFSO.CreateFolder(strDestFolderPath & "\" & strDestFolderName & "")
Loop Until objFSO.FolderExists(strDestFolderPath & "\" & strDestFolderName & "") <> 0

'Criando link target.lnk na pasta " & strDestFolderName & "
set objShortcut = WshShell.CreateShortcut(strDestFolderPath & "\" & strDestFolderName & "\target.lnk")
	objShortcut.FullName = "target.lnk"
	objShortcut.TargetPath = strTargetLink
	objShortcut.Save
	
'Criando arquivo desktop.ini
If objFSO.FileExists(strDestFolderPath & "\" & strDestFolderName & "\desktop.ini") Then
	Set objDesktopFile = objFSO.OpenTextFile(strDestFolderPath & "\" & strDestFolderName & "\desktop.ini", ForWriting, True)
Else
	Set objDesktopFile = objFSO.CreateTextFile(strDestFolderPath & "\" & strDestFolderName & "\desktop.ini")
End If

'Criar desktop.ini
strDeskIni = "[.ShellClassInfo]" & vbCr & vbCrlf _
		   & "CLSID2={0AFACED1-E828-11D1-9187-B532F1E9575D}"  & vbCr & vbCrlf _
		   & "Flags=2"  & vbCr & vbCrlf _
		   & "ConfirmFileOp=0"

objDesktopFile.Writeline strDeskIni
objDesktopFile.Close

' Definir que a pasta " & strDestFolderName & " terá atributos de sistema
WshShell.Run "C:\windows\system32\attrib.exe +s " & strDestFolderPath & "\" & strDestFolderName & "", 0, TRUE
' Set on system flag to the shortcut folder
MsgBox "Set security properties for the folder shortcut and click OK", vbOKOnly, strTitle
objFolder.Attributes = objFolder.Attributes + Attrib_ReadOnly + Attrib_System
Wscript.Quit