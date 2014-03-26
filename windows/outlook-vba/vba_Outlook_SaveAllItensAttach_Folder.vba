Sub Save_MSO_ATTACHMENTS()
    
    On Error GoTo Err_OL
    Dim ns As NameSpace
    Dim Inbox As MAPIFolder
    Dim Item As Object
    Dim Atmt As Attachment
    Dim FileName As String
    Dim i As Integer

	Set ns = GetNamespace("MAPI")
    'Set Inbox = ns.GetDefaultFolder(olFolderInbox) 'Set folder manualy
    Set Inbox = Application.ActiveExplorer.CurrentFolder ' Set folder to Current Folder
    i = 0
    sSaveAttachmentsFolder = "UNIDADE:\CAMINHO\"
     
    If Inbox.Items.Count = 0 Then
        MsgBox "There are no messages in the Inbox.", vbInformation, _
               "Nothing Found"
        Exit Sub
    End If
     
    For Each Item In Inbox.Items
        sAttachName = Split(Item.Subject, " ", -1, vbTextCompare)
        For Each Atmt In Item.Attachments
           sAttachExt = Split(Atmt.FileName, ".", -1, vbTextCompare)
           FileName = sSaveAttachmentsFolder & "oAttach_" & sAttachName(1) & "." & sAttachExt(1)
           Atmt.SaveAsFile FileName
           i = i + 1
        Next Atmt
    Next Item
    
Err_OL:
    If Err <> 0 Then
    MsgBox Err.Number & " - " & Err.Description
    Err.Clear
    Resume Next
    End If
    Exit Sub

End Sub