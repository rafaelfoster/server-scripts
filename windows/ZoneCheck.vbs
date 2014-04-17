On Error Resume Next
Set WshShell = CreateObject("Wscript.Shell")
set WshEnv = WshShell.Environment("PROCESS")

' Disable Zone Check
WshEnv("SEE_MASK_NOZONECHECKS") = 1

' Enable Zone Check again
WshEnv.Remove("SEE_MASK_NOZONECHECKS")