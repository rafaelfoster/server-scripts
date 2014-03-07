On Error Resume Next
Set WshShell = CreateObject("Wscript.Shell")
set WshEnv = WshShell.Environment("PROCESS")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set SystemSet = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 
Const Log_Anexar = 2 '( 1 = Read, 2 = Write, 8 = Append )

' Disable Zone Check
WshEnv("SEE_MASK_NOZONECHECKS") = 1

' Enable Zone Check again
WshEnv.Remove("SEE_MASK_NOZONECHECKS")