On Error Resume Next

Const ForReading = 1
Const ForWriting = 2
Const ForAppend = 8
Const OverwriteExisting = TRUE
Const Attrib_ReadOnly = 1
Const Attrib_System = 4

' --------------------------------------------------------------------------------------------
' Create instance to handle filesystem files
Set WshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

strTitle = "Create folder shortcut"
strRealShortcutName = "target.lnk"
strIniName = "desktop.ini"
strIniContent = "[.ShellClassInfo]" & vbCrlf _
    & "CLSID2={0AFACED1-E828-11D1-9187-B532F1E9575D}"  & vbCrlf _
    & "Flags=2"  & vbCrlf _
    & "ConfirmFileOp=0"
bQuit = False
strShortcutPath = ""
strShortcutName = ""
strTarget = ""

Do
    bQuit = False
    GetUserInput
    CreateShortcut
    
    If MsgBox("Do you want to create another shortcut?", vbYesNo, strTitle) = vbYes Then
        bQuit = True
    End If
Loop Until Not(bQuit)

Sub GetUserInput
    strShortcutPath = InputBox("Type the shortcut location", strTitle, strShortcutPath)
    If strShortcutPath = "" Then WScript.Quit(1)

    strShortcutName = InputBox("Type the shortcut name", strTitle, strShortcutName)
    If strShortcutName = "" Then WScript.Quit(1)

    strTarget = InputBox("Type the target location", strTitle, strTarget)
    If strTarget = "" Then WScript.Quit(1)

    strMsgAck = "A folder shortcut will be created with following properties:" & vbCrLf _
        & "Location: " & strShortcutPath & vbCrLf _
        & "Name: " & strShortcutName & vbCrLf _
        & "Target: " & strTarget & vbCrLf & vbCrLf _
        & "Are you sure you want to create this shortcut?"
    If MsgBox(strMsgAck, vbYesNo, strTitle) = vbNo Then
        WScript.Quit(1)
    End If
End Sub

Sub CreateShortcut
    ' Create a folder to become the shortcut
    strShortcut = strShortcutPath & "\" & strShortcutName
    Set objFolder = objFSO.CreateFolder(strShortcut)
    If Not(objFSO.FolderExists(strShortcut)) Then
        MsgBox "Can't create a folder at " & strShortcutPath, vbOKOnly, strTitle
        WScript.Quit(1)
    End If

    ' Create a shortcut inside the folder
    strRealShortcut = strShortcut & "\" & strRealShortcutName
    Set objShortcut = WshShell.CreateShortcut(strRealShortcut)
    'objShortcut.FullName = strRealShortcutName
    objShortcut.TargetPath = strTarget
    objShortcut.Save

    ' Create desktop.ini file
    strIni = strShortcut & "\" & strIniName
    Set objDesktopFile = objFSO.OpenTextFile(strIni, ForWriting, True)
    objDesktopFile.Writeline(strIniContent)
    objDesktopFile.Close

    ' Set on system flag to the shortcut folder
    MsgBox "Set security properties for the folder shortcut and click OK", vbOKOnly, strTitle
    objFolder.Attributes = objFolder.Attributes + Attrib_ReadOnly + Attrib_System
End Sub

