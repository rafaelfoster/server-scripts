On Error Resume Next
Set WshShell = CreateObject("Wscript.Shell")
set WshEnv = WshShell.Environment("PROCESS")
Set objFSO       = CreateObject("Scripting.FileSystemObject")
Set SystemSet = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 

Const ForReading = 1
Const ForWriting = 2
Const ForAppend  = 8
Const OverwriteExisting = TRUE

WshEnv("SEE_MASK_NOZONECHECKS") = 1
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
LogFile = "\\rodrimar.com.br\TI\AD-MGT\Logs\Log_Patch_GPO_WindowsXP\updateLog_" & strUserName & ".log"
For each System in SystemSet 
	SysOperation = System.Caption & " SP" & System.ServicePackMajorVersion & " " & System.BuildNumber
Next

If objFSO.FileExists(LogFile) Then
	Set objCriaLog = objFSO.OpenTextFile(LogFile, ForReading, True)
Else
	Set objCriaLog = objFSO.CreateTextFile(LogFile)
End If

objCriaLog.WriteLine "Data: " & date & " - " & time

'If ( InStr(LCase(SysOperation),"xp" ) <> 0 ) Then
'	Result1 = WshShell.Run("\\rodrimar.com.br\netlogon\VBScripts\Files\ext_gpo_KB943729.exe /norestart /quiet",0,true)
'	objCriaLog.WriteLine "Atualizacao KB943729 instalada com exito"
'End If

objCriaLog.WriteLine "--- Fim do GPUPDATE ---"

Result = WshShell.Run("cmd /c echo n | gpupdate /force",0,true)
WshEnv.Remove("SEE_MASK_NOZONECHECKS")
Wscript.Quit(Result)


'------ GPRESULT --------------------------------------------------------------

'Aguardar 10 minutos antes de iniciar
Wscript.Sleep 600000

LogFile = "\\rodrimar.com.br\TI\AD-MGT\Logs\Log_GPUpdate\gpresult_" & strUserName

If ( InStr(LCase(SysOperation),"xp" ) <> 0 ) Then
	
	LogFile = LogFile & ".log"
	WshShell.Run "cmd /c gpresult > " & LogFile,0,TRUE
	
	'Do
	'	WScript.StdOut.WriteLine(strGPResult.StdOut.ReadLine())
	'	objCriaLog.Writeline strGPResult.StdOut.ReadLine()
	'Loop While Not strGPResult.Stdout.atEndOfStream
	'WScript.StdOut.WriteLine(strGPResult.StdOut.ReadAll)
	
	ObjCriaLog.Close
	
Else

	LogFile = LogFile & ".html"
	WshShell.Run "gpresult /F /H " & LogFile,0,TRUE
	
End If
