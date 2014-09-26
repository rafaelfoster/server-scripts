On Error Resume Next
Set objArgs = Wscript.Arguments
Set WshShell = CreateObject("WScript.Shell")
Set objRegEx = CreateObject("VBScript.RegExp")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set SystemSet    = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 

Const ForReading = 1
Const ForWriting = 2
Const ForAppend  = 8
Const OverwriteExisting = TRUE

strComputer = "."
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
strUserPath = wshShell.ExpandEnvironmentStrings( "%USERPROFILE%" )
strSessionName = wshShell.ExpandEnvironmentStrings( "%SESSIONNAME%" )
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )

strRepoBase = "\\rodrimar.com.br\ti\AD-MGT\scripts\repositories\certificados"
Log_Folder = "\\rodrimar.com.br\Ti\AD-MGT\Logs\Log_InstalacaoCertificado"
Log_File = Log_Folder & "\Log_Certificado_" & strUserName & ".txt"

'Aguardar 1 minuto antes de iniciar
'Wscript.Sleep 60000

For each System in SystemSet 
	SysVersion     = System.Caption & " SP" & System.ServicePackMajorVersion & " " & System.BuildNumber
	dtmConvertedDate.Value = System.InstallDate
    dtmInstallDate = dtmConvertedDate.GetVarDate
	SysInstallDate = dtmInstallDate
Next

if ( inStr(LCase(SysVersion),"server") = 0 ) Then
	Wscript.Quit
End If

If WScript.Arguments.Count = 0 then
    WScript.Echo  "This script require arguments" & vbCr _
				& "**************************************************************************************" & vbCr _
				& "Run this with: " & Wscript.ScriptFullName & " certificate_path\certificate_pfx " & "certificate_password" & vbCr & vbCr _
				& "Where [certificate_path] should be in default repo folder (" & strRepoBase & ")" & vbCr _
				& "Example: " & vbCr & vbCr _
				& Wscript.ScriptFullName & " my_certificate\certificate.pfx " & "V3ryS3cUreP4ssw0rd" & vbCr 
	If InStr(1, WScript.FullName, "cscript", vbTextCompare) Then
		Wscript.Echo "saindo..."
		Wscript.Quit
	ElseIf InStr(1, WScript.FullName, "wscript", vbTextCompare) Then
		strCertPath = InputBox("Digite o caminho do Certificado " & vbCr & "Seguindo o padrao: certificate_path\certificate_pfx")
		strCertPass = InputBox("Digite a senha para instalar o Certificado")
	End If
Else 
	If ( WScript.Arguments.Count <> 0 )  Then
		strCertPath = objArgs(0)
	End If
	'If ( WScript.Arguments.Count <> 1 ) Then
	'	strCertPass = objArgs(1)
	'End If
End if

If( IsEmpty(strCertPath) ) Then
	Wscript.Echo "Parametros de caminho não informados. " & vbCr & "Encerrando...."
End If

' ------------------------------------------------------------------------------------------------------------------------------------------------
' Inicia a geração do Arquivo de Log
If objFSO.FileExists(Log_File) Then
	Set objCriaLog = objFSO.OpenTextFile(Log_File, ForAppend, True)
Else
	Set objCriaLog = objFSO.CreateTextFile(Log_File)
End If

'Certificado .PFX
'InstallCommand = "certutil -f -user -p """ & strCertPass & """ -installCert " & strRepoBase & "\" & strCertPath

'Certificado .CER
InstallCommand = "certutil -addstore -user -f ""My"" " & strRepoBase & "\" & strCertPath

Result = WshShell.Run(InstallCommand,0,FALSE)

Wscript.Sleep 500
WshShell.AppActivate "Security Warning"
WshShell.SendKeys "Y"
WshShell.SendKeys "S"

objCriaLog.WriteLine 
objCriaLog.WriteLine "Data: " & date & " - " & time & " - Computador: " & strComputerName & " | Usuario: " & strUserName
objCriaLog.WriteLine "Certificado instalado: " & strRepoBase & "\" & strCertPath