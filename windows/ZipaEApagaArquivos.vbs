On Error Resume Next
Set WshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

objStartFolder = Replace(WScript.ScriptFullName, WScript.ScriptName, "")

str_TransferFolder = objStartFolder & "PASTA1\"
str_OLDransferFolder = objStartFolder & "PASTA2\"

WshShell.CurrentDirectory = str_TransferFolder

Set objFolder = objFSO.GetFolder(str_TransferFolder)
Set colFiles = objFolder.Files

For Each objFile in colFiles

	strTmpFolderName = split(objFile.DateLastModified ,"/")
	strFolderDateName = split(strTmpFolderName(2)," ")(0) & "-" & strTmpFolderName(1) & "-" & strTmpFolderName(0)
	If objFSO.FolderExists( strFolderDateName ) = False Then
		objFSO.CreateFolder( strFolderDateName  )
	End If
	objFSO.MoveFile objFile.Name, strFolderDateName & "\" & objFile.Name

Next

'WshShell.CurrentDirectory = ".."

Set objFolder2 = objFSO.GetFolder(str_TransferFolder)
Set colFiles1 = objFolder2.SubFolders

For Each objFile in colFiles1

	zipfile = objFile.Name & ".zip"

	ArchiveFolder zipfile, objFile.Name

	objFSO.DeleteFolder objFile.Name
	objFSO.MoveFile zipfile , str_OLDransferFolder

Next

Sub ArchiveFolder (zipFile, sFolder)

    With CreateObject("Scripting.FileSystemObject")
        zipFile = .GetAbsolutePathName(zipFile)
        sFolder = .GetAbsolutePathName(sFolder)

        With .CreateTextFile(zipFile, True)
            .Write Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, chr(0))
        End With
    End With

    With CreateObject("Shell.Application")
        .NameSpace(zipFile).CopyHere .NameSpace(sFolder).Items

        Do Until .NameSpace(zipFile).Items.Count = _
                 .NameSpace(sFolder).Items.Count
            WScript.Sleep 1000 
        Loop
    End With

End Sub