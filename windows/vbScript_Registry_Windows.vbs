On Error Resume Next
' --------------------------------------------------------------------------------------------
' Criação do Objeto para criação/gravação do arquivo
set WshShell     = CreateObject("WScript.Shell")
Set WshNetwork   = CreateObject("WScript.Network")
Set objFSO       = CreateObject("Scripting.FileSystemObject")
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")

Const HKEY_LOCAL_MACHINE = &H80000002
Const ForReading = 1
Const ForWriting = 2
Const ForAppend  = 8
Const OverwriteExisting = TRUE

' --------------------------------------------------------------------------------------------------------------------------------
' Realizar a leitura de uma chave de registro

If objFSO.FileExists("ARQUIVO_DE_TEXTO") = False Then
	Set objCriaLog = objFSO.CreateTextFile("ARQUIVO_DE_TEXTO")
	objCriaLog.WriteLine "ComputerName;UserName;IPAddress;CurrentGroup;SymantecRegKey"
	objCriaLog.WriteLine
	objCriaLog.Close
End If

strKeyStatus = "HKLM\SOFTWARE\Symantec\Symantec Endpoint Protection\SMC\SYLINK\SyLink\CommunicationStatus"
SymantecStatus = WshShell.RegRead(strKeyStatus)