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

Wscript.Quit(strExecFile)
