' Script para instalar certificados no Firefox
' Baixe o utilitário certutil para firefox em:
' https://www.dropbox.com/s/1mwpmewsp92pg7h/firefox_certutil.zip?dl=0


On Error Resume Next

'Aguardar 1 minuto antes de iniciar
Wscript.Sleep 60000

Set objArgs   = Wscript.Arguments
Set WshShell  = CreateObject("WScript.Shell")
Set objRegEx  = CreateObject("VBScript.RegExp")
Set objFSO    = CreateObject("Scripting.objFSOObject")
Set SystemSet = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem")

Const ForReading = 1
Const ForWriting = 2
Const ForAppend  = 8
Const OverwriteExisting = TRUE

strComputer     = "."
strUserName     = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
strUserPath     = wshShell.ExpandEnvironmentStrings( "%USERPROFILE%" )
strSessionName  = wshShell.ExpandEnvironmentStrings( "%SESSIONNAME%" )
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
strProfileLocation = strUserPath & "\AppData\Roaming\Mozilla\Firefox\profiles.ini"
strDBFileLocation = "cert8.db"

'----------------------------------
' Variaveis para alterar
Log_Folder    = strUserPath
Log_File_Name = "Log_Certificado_" & strUserName & ".txt"

CertutilPath      = "C:\Users\Rafael Foster\Desktop\firefox_add-certs\firefox_add-certs\bin"
certFile          = "C:\Users\Rafael Foster\Desktop\Fortinet_CA_SSL.cer"
certName          = "Fortinet SSL CA"

'--------------------------------------------
' Pular a instalação em Servidores
For each System in SystemSet
	SysVersion     = System.Caption & " SP" & System.ServicePackMajorVersion & " " & System.BuildNumber
	dtmConvertedDate.Value = System.InstallDate
    dtmInstallDate = dtmConvertedDate.GetVarDate
	SysInstallDate = dtmInstallDate
Next

if ( inStr(LCase(SysVersion),"server") = 0 ) Then
	Wscript.Quit
End If
'---------------------------------------------

Log_File = Log_Folder & "\" & Log_File_Name

' ------------------------------------------------------------------------------------------------------------
' Inicia a geração do Arquivo de Log
If objFSO.FileExists(Log_File) Then
	Set objCriaLog = objFSO.OpenTextFile(Log_File, ForAppend, True)
Else
	Set objCriaLog = objFSO.CreateTextFile(Log_File)
End If


If (objFSO.FileExists(strProfileLocation)) Then
  strData = objFSO.OpenTextFile(strProfileLocation ,ForReading).ReadAll
  arrLines = Split(strData,vbCrLf)
  For Each strLine in arrLines
    If Left(strLine, 14) = "Path=Profiles/" then
      strProfileName = Right(strLine, (len(strLine) - 14))
    End if
  Next

  strProfileFolder = strUserPath & "\AppData\Roaming\Mozilla\Firefox\Profiles\" & strProfileName

  if (objFSO.FolderExists(strProfileFolder)) Then
    certDBFile = strProfileFolder & "\cert8.db"
    oldfile = strProfileFolder & "\cert8.old"
    objFSO.CopyFile certDBFile, oldfile, True
  End if

  InstallCommand = CertutilPath & "\certutil.exe -A -d " & strProfileFolder & " -i certFile -n " & certName & " -t CT,c,c"

  Result = WshShell.Run(InstallCommand,0,FALSE)

  objCriaLog.WriteLine
  objCriaLog.WriteLine "Data: " & date & " - " & time & " - Computador: " & strComputerName & " | Usuario: " & strUserName
  objCriaLog.WriteLine "Certificado instalado: " & strRepoBase & "\" & strCertPath

Else
  objCriaLog.WriteLine
  objCriaLog.WriteLine "Data: " & date & " - " & time & " - Computador: " & strComputerName & " | Usuario: " & strUserName
  objCriaLog.WriteLine "Certificado não instalado!"
  objCriaLog.WriteLine "Não foi possivel encotrar o perfil do Firefox."

  Wscript.Quit()
End if
