'Script para demonstrar e desabilitar o Aviso de Segurança ao executar um binario via VBS
On Error Resume Next
Set WshShell = CreateObject("Wscript.Shell")
set WshEnv = WshShell.Environment("PROCESS")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set SystemSet = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 
Const Log_Anexar = 2 '( 1 = Read, 2 = Write, 8 = Append )

WshEnv("SEE_MASK_NOZONECHECKS") = 1
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )

strExecFile = WshShell.Run("KB943729.exe /norestart /quiet",0,true)

WshEnv.Remove("SEE_MASK_NOZONECHECKS")

'Wscript.Quit(strExecFile)

'------ Gerar report do GPRESULT --------------------------------------------------------------

'Aguardar 10 minutos antes de iniciar
Wscript.Sleep 600000

LogFile = "\\SERVIDOR\PASTA\gpresult_" & strUserName

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
